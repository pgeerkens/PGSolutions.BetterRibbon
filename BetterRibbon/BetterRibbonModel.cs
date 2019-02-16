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
    public sealed class BetterRibbonModel : IRibbonDispatcher {
        internal BetterRibbonModel(BetterRibbonViewModel viewModel) {
            ViewModel   = viewModel;

            BrandingModel        = new BrandingModel(ViewModel?.BrandingViewModel, BrandingIcon);
            LinksAnalysisModel   = new LinksAnalysisModel(ViewModel?.LinksAnalysisViewModel);
            VbaSourceExportModel = new VbaSourceExportModel(
                    new List<VbaSourceExportGroupModel>() {
                        new VbaSourceExportGroupModel(ViewModel?.VbaExportViewModel_MS,"MS"),
                        new VbaSourceExportGroupModel(ViewModel?.VbaExportViewModel_PG,"PG")
                    });
            CustomButtonsModel   = new CustomButtonsModel(ViewModel.CustomButtonsViewModel);
        }

        private  BetterRibbonViewModel ViewModel            { get; }

        internal BrandingModel         BrandingModel        { get; private set; }
        internal LinksAnalysisModel    LinksAnalysisModel   { get; private set; }
        internal VbaSourceExportModel  VbaSourceExportModel { get; private set; }
        internal CustomButtonsModel    CustomButtonsModel   { get; private set; }


        #region IRibbonDispatcher methods
         /// <inheritdoc/>
        public void Invalidate() {
            BrandingModel?.Invalidate();
            LinksAnalysisModel?.Invalidate();
            VbaSourceExportModel?.Invalidate();
            CustomButtonsModel?.Invalidate();
        }

         /// <inheritdoc/>
        public void InvalidateCustomControlsGroup() => CustomButtonsModel?.Invalidate();

         /// <inheritdoc/>
        public void InvalidateControl(string ControlId) => ViewModel?.InvalidateControl(ControlId);

        /// <inheritdoc/>
        public void DetachProxy(string controlId)
        => CustomButtonsModel.GetControl<IRibbonCommon>(controlId)?.Detach();

        /// <inheritdoc/>
        [SuppressMessage( "Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed",
                Justification = "Matches COM usage." )]
        public IRibbonControlStrings NewControlStrings(string label,
                string screenTip = null, string superTip = null,
                string keyTip = null, string alternateLabel = null, string description = null) =>
            ViewModel.RibbonFactory.NewControlStrings(label, screenTip,
                    superTip, keyTip, alternateLabel, description);
        #endregion

        private static IPictureDisp BrandingIcon => Resources.PGeerkens.ImageToPictureDisp();

        /// <inheritdoc/>
        public ISelectableItemModel NewSelectableModel(string controlID, IRibbonControlStrings strings) {
            var vm = ViewModel.RibbonFactory.NewSelectableItem(controlID);
            var model = new SelectableItemModel(id => vm, strings, true, true)
                        .Attach(controlID);
            return model;
        }

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IRibbonButtonModel NewRibbonButtonModel(IRibbonControlStrings strings,
                IPictureDisp image = null, bool isEnabled = true, bool isVisible = true)
        => CustomButtonsModel.NewButtonModel(strings, image, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IRibbonButtonModel NewRibbonButtonModelMso(IRibbonControlStrings strings,
                string imageMso = "MacroSecurity", bool isEnabled = true, bool isVisible = true)
        => CustomButtonsModel.NewButtonModel(strings, imageMso, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IRibbonToggleModel NewRibbonToggleModel(IRibbonControlStrings strings,
                IPictureDisp image = null, bool isEnabled = true, bool isVisible = true)
        => CustomButtonsModel.NewToggleModel(strings, image, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IRibbonToggleModel NewRibbonToggleModelMso(IRibbonControlStrings strings,
                string imageMso = "MacroSecurity", bool isEnabled = true, bool isVisible = true)
        => CustomButtonsModel.NewToggleModel(strings, imageMso, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IRibbonDropDownModel NewRibbonDropDownModel(IRibbonControlStrings strings,
                bool isEnabled = true, bool isVisible = true)
        => CustomButtonsModel.NewDropDownModel(strings, isEnabled, isVisible);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IRibbonGroupModel NewRibbonGroupModel(IRibbonControlStrings strings,
                bool isEnabled = true, bool isVisible = true)
        => new RibbonGroupModel(id => CustomButtonsModel.GetControl<RibbonGroupViewModel>(id),
                strings, isEnabled, isVisible, CustomButtonsModel);
    }
}

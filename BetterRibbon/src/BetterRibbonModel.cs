////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.BetterRibbon.VbaSourceExport;
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
    [Serializable, CLSCompliant(false)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IRibbonDispatcher))]
    [Guid(RibbonDispatcher.Guids.BetterRibbon)]
    [ProgId(ProgIds.RibbonDispatcherProgId)]
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
        Justification = "Public, Non-Creatable, class with exported Events.")]
    //[Guid("A8ED8DFB-C422-4F03-93BF-FB5453D8F213")]
    public sealed class BetterRibbonModel : IRibbonDispatcher {
        internal BetterRibbonModel(BetterRibbonViewModel viewModel) {
            ViewModel   = viewModel;

            if(viewModel.IsInitialized) {
                OnViewModelInitialized();
            } else {
                ViewModel.Initialized += ViewModelInitialized;
            }
        }

        private void ViewModelInitialized(object sender, EventArgs e) {
            OnViewModelInitialized();
            ViewModel.Initialized -= ViewModelInitialized;
        }

        private void OnViewModelInitialized() {
            BrandingModel        = new BrandingModel(ViewModel.BrandingViewModel);
            LinksAnalysisModel   = new LinksAnalysisModel(ViewModel.LinksAnalysisViewModel);
            VbaSourceExportModel = new VbaSourceExportModel(
                new List<IVbaSourceExportViewModel> {
                    ViewModel.VbaExportViewModel_MS,
                    ViewModel.VbaExportViewModel_PG
                } );

            CustomButtonsModel   = new CustomButtonsModel(ViewModel.CustomButtonsViewModel);
            DemonstrationModel   = new DemonstrationModel(ViewModel.DemonstrationViewModel);
            Attach();
        }

        private BetterRibbonViewModel ViewModel            { get; set; }

        private BrandingModel         BrandingModel        { get; set; }
        private LinksAnalysisModel    LinksAnalysisModel   { get; set; }
        [SuppressMessage( "Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode" )]
        private VbaSourceExportModel  VbaSourceExportModel { get; set; }
        private DemonstrationModel    DemonstrationModel   { get; set; }
        private CustomButtonsModel    CustomButtonsModel   { get; set; }

         /// <inheritdoc/>
        public void Attach() {
            BrandingModel.Attach();
            DemonstrationModel.Attach();
        }
         /// <inheritdoc/>
        public void Detach() {
            DemonstrationModel.Detach();
            BrandingModel.Detach();
        }

        #region IRibbonDispatcher methods
         /// <inheritdoc/>
        public void Invalidate() {
            BrandingModel?.Invalidate();
            DemonstrationModel?.Invalidate();
            CustomButtonsModel?.Invalidate();
        }

         /// <inheritdoc/>
        public void InvalidateCustomControlsGroup() => CustomButtonsModel?.Invalidate();

         /// <inheritdoc/>
        public void InvalidateControl(string ControlId) => ViewModel?.InvalidateControl(ControlId);

         /// <inheritdoc/>
        public IRibbonButton   AttachButton(string controlId, IRibbonControlStrings strings) =>
            SetStrings(GetControl<RibbonButton>(controlId),strings).Attach() as IRibbonButton;

         /// <inheritdoc/>
        public IRibbonToggle   AttachCheckBox(string controlId, IRibbonControlStrings strings,
                IBooleanSource source) =>
            SetStrings(GetControl<RibbonCheckBox>(controlId),strings).Attach(source.Getter);

         /// <inheritdoc/>
        public IRibbonDropDown AttachDropDown(string controlId, IRibbonControlStrings strings,
                IIntegerSource source) =>
            SetStrings(GetControl<RibbonDropDown>(controlId),strings).Attach(source.Getter);

         /// <inheritdoc/>
        public IRibbonToggle   AttachToggle(string controlId, IRibbonControlStrings strings,
                IBooleanSource source) =>
            SetStrings(GetControl<RibbonToggleButton>(controlId),strings).Attach(source.Getter);

        /// <inheritdoc/>
        public void DetachProxy(string controlId) =>
            CustomButtonsModel.GetControl<RibbonCommon>(controlId)?.Detach();

        /// <inheritdoc/>
        [SuppressMessage( "Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed",
                Justification = "Matches COM usage." )]
        public IRibbonControlStrings NewControlStrings(string label,
                string screenTip = "", string superTip = "",
                string keyTip    = "", string alternateLabel = "", string description = "") =>
            ViewModel.RibbonFactory.NewControlStrings(label,
                    screenTip, superTip, keyTip, alternateLabel, description);

        /// <inheritdoc/>
        public void ShowInactive(bool showWhenInactive) {
            CustomButtonsModel.SetShowWhenInactive(showWhenInactive);
            InvalidateCustomControlsGroup();
        }
        #endregion

        private TControl GetControl<TControl>(string controlId) where TControl:RibbonCommon =>
            CustomButtonsModel.GetControl<TControl>(controlId);

        private static TControl SetStrings<TControl>(TControl ctrl, IRibbonControlStrings strings) where TControl:RibbonCommon {
            ctrl?.SetLanguageStrings(strings ?? RibbonControlStrings.Default(ctrl.Id));
            return ctrl;
        }
    }
}

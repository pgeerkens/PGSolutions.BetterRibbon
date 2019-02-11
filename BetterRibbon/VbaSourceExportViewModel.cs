////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.BetterRibbon;
using PGSolutions.RibbonDispatcher.Utilities;

namespace PGSolutions.RibbonUtilities.VbaSourceExport {
    using VbaExportEventHandler = EventHandler<VbaExportEventArgs>;

    /// <summary>.</summary>
    [CLSCompliant(false)]
    public sealed class VbaSourceExportViewModel : AbstractRibbonGroupViewModel, IVbaSourceExportViewModel, IApplication {
        /// <summary>.</summary>
        public VbaSourceExportViewModel(IRibbonFactory factory, string suffix) : base(factory) {
            var defaultSize = suffix=="MS" ? false : true;
            VbASourceExportGroup  = Factory.NewRibbonGroup($"VbaExportGroup{suffix}");

            UseSrcFolderToggle    = Factory.NewRibbonToggleMso($"UseSrcFolderToggle{suffix}",
                                            isLarge:defaultSize, imageMso:ToggleImage(false));
            SelectedProjectButton = Factory.NewRibbonButtonMso($"SelectedProjectButton{suffix}",
                                            isLarge:defaultSize, imageMso:"SaveAll", showImage:true);
            CurrentProjectButton  = Factory.NewRibbonButtonMso($"CurrentProjectButton{suffix}",
                                            isLarge:defaultSize, imageMso:"FileSaveAs", showImage:true);

            UseSrcFolderToggle.Toggled    += OnSrcFolderToggled;
            SelectedProjectButton.Attach<RibbonButton>().Clicked += OnExportSelected;
            CurrentProjectButton.Attach<RibbonButton>().Clicked  += OnExportCurrent;
        }

        /// <inheritdoc/>
        public void Attach(IBooleanSource srcToggleSource) =>
            UseSrcFolderToggle.Attach(srcToggleSource.Getter);

        /// <inheritdoc/>
        public void Invalidate() => UseSrcFolderToggle.Invalidate();

        /// <inheritdoc/>
        public event ToggledEventHandler   UseSrcFolderToggled;
        /// <inheritdoc/>
        public event VbaExportEventHandler SelectedProjectsClicked;
        /// <inheritdoc/>
        public event VbaExportEventHandler CurrentProjectClicked;

        private void OnSrcFolderToggled(object sender, bool isPressed) {
            UseSrcFolderToggle.SetImageMso(ToggleImage(isPressed));
            UseSrcFolderToggled?.Invoke(sender, isPressed);
        }

        private static string ToggleImage(bool isPressed)
        => isPressed ? "TagMarkComplete" : "MarginsShowHide";

        private void OnExportCurrent(object sender) {
            if (!IsProjectModelTrusted()) { return; }
            PerformSilently(
                () => CurrentProjectClicked?.Invoke(this, new VbaExportEventArgs(new ProjectFilterExcel(this))
            ));
        }

        private void OnExportSelected(object sender) {
            if (!IsProjectModelTrusted()) { return; }
            var securitySaved = Application.AutomationSecurity;
            Application.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;

            try {
                var fd = Application.FileDialog[MsoFileDialogType.msoFileDialogFilePicker];
                fd.Title = "Select VBA Project(s) to Export From";
                fd.ButtonName = "Export";
                fd.AllowMultiSelect = true;
                fd.Filters.Clear();
                fd.InitialFileName = Application.ActiveWorkbook?.Path ?? "C:\\";

                var list = new ProjectFilters(this);
                foreach (var item in list) {
                    fd.Filters.Add(item.Description, item.Extensions);
                }
                if (fd.Show() != 0) {
                    PerformSilently(
                        () => SelectedProjectsClicked?.Invoke(this,
                                new VbaExportEventArgs(list[fd.FilterIndex-1], fd.SelectedItems)
                    ));
                }
            }
            finally {
                Application.AutomationSecurity = securitySaved;
            }
        }

        private bool IsProjectModelTrusted() {
            try {
                return Application.VBE != null;
            }
            catch (COMException) { PleaseEnableTrust(); }
            catch (InvalidOperationException) { PleaseEnableTrust(); }
            return false;
        }

        private static void PerformSilently(System.Action action) {
            try {
                Application.Cursor = XlMousePointer.xlWait;

                action();
            } finally {
                Application.StatusBar = false;

                Application.Cursor = XlMousePointer.xlDefault;
            }
        }

        private static void PleaseEnableTrust()
        => "Please enable trust of the Project Object Model".MsgBoxShow("Project Model Not Trusted");

        static Application Application => Globals.ThisAddIn.Application;

        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        private  RibbonGroup        VbASourceExportGroup  { get; }
        private  RibbonToggleButton UseSrcFolderToggle    { get; }
        /// <inheritdoc/>
        public   RibbonButton       SelectedProjectButton { get; }
        /// <inheritdoc/>
        public   RibbonButton       CurrentProjectButton  { get; }

        IRibbonToggle IVbaSourceExportViewModel.UseSrcFolderToggle    => UseSrcFolderToggle;
        IRibbonButton IVbaSourceExportViewModel.SelectedProjectButton => SelectedProjectButton;
        IRibbonButton IVbaSourceExportViewModel.CurrentProjectButton  => CurrentProjectButton;

        /// <inheritdoc/>
        public Workbook ActiveWorkbook => Application.ActiveWorkbook;

        /// <inheritdoc/>
        public bool     DisplayAlerts {
            get => Application.DisplayAlerts;
            set => Application.DisplayAlerts = value;
        }

        /// <inheritdoc/>
        public dynamic  StatusBar {
            get => Application.StatusBar;
            set => Application.StatusBar = value;
        }

        /// <inheritdoc/>
        public MsoAutomationSecurity AutomationSecurity {
            get => Application.AutomationSecurity;
            set => Application.AutomationSecurity = value;
        }
    }
}

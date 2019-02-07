﻿////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Windows.Forms;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using static Microsoft.Office.Core.RibbonControlSize;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.BetterRibbon;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonUtilities.VbaSourceExport {
    using VbaExportEventHandler = EventHandler<VbaExportEventArgs>;

    /// <summary>.</summary>
    [CLSCompliant(false)]
    public class VbaSourceExportViewModel : AbstractRibbonGroupViewModel, IVbaSourceExportViewModel, IApplication {
        public VbaSourceExportViewModel(IRibbonFactory factory, string suffix) : base(factory) {
            var defaultSize = suffix=="MS" ? RibbonControlSizeRegular : RibbonControlSizeLarge;
            VbASourceExportGroup  = Factory.NewRibbonGroup($"VbaExportGroup{suffix}");

            UseSrcFolderToggle    = Factory.NewRibbonToggleMso($"UseSrcFolderToggle{suffix}",
                                            size:defaultSize, imageMso:"MacroSecurity");
            SelectedProjectButton = Factory.NewRibbonButtonMso($"SelectedProjectButton{suffix}",
                                            size:defaultSize, imageMso:"RefreshAll", showImage:true);
            CurrentProjectButton  = Factory.NewRibbonButtonMso($"CurrentProjectButton{suffix}",
                                            size:defaultSize, imageMso:"Refresh", showImage:true);

            UseSrcFolderToggle.Toggled    += OnSrcFolderToggled;
            SelectedProjectButton.Attach<RibbonButton>().Clicked += OnExportSelected;
            CurrentProjectButton.Attach<RibbonButton>().Clicked  += OnExportCurrent;
        }

        public void Attach(IBooleanSource srcToggleSource) =>
            UseSrcFolderToggle.Attach(srcToggleSource.Getter);

        public void Invalidate() => UseSrcFolderToggle.Invalidate();

        public event ToggledEventHandler   UseSrcFolderToggled;
        public event VbaExportEventHandler SelectedProjectsClicked;
        public event VbaExportEventHandler CurrentProjectClicked;

        private void OnSrcFolderToggled(object sender, bool isPressed) =>
            UseSrcFolderToggled?.Invoke(sender,isPressed);

        private void OnExportCurrent(object sender) {
            try {
                if ( Application.VBE == null) { throw new COMException(); }
                PerformSilently(
                    () => CurrentProjectClicked?.Invoke(this, new VbaExportEventArgs(new ProjectFilterExcel(this) )
                ) );
            } catch (COMException) { PleaseEnableTrust(); }
        }

        private void OnExportSelected(object sender) {
            try {
                if ( Application.VBE == null) { throw new COMException(); }
                var securitySaved = Application.AutomationSecurity;
                Application.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;

                try {
                    var fd = Application.FileDialog[MsoFileDialogType.msoFileDialogFilePicker];
                    fd.Title = "Select VBA Project(s) to Export From";
                    fd.ButtonName = "Export";
                    fd.AllowMultiSelect = true;     // MultiSelect requires eponymous naming
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
                        ) );
                    }
                } finally {
                    Application.DisplayAlerts = true;
                    Application.ScreenUpdating = true;
                    Application.Cursor = XlMousePointer.xlDefault;
                    Application.AutomationSecurity = securitySaved;
                }
            } catch (COMException) { PleaseEnableTrust(); }
        }

        private void PleaseEnableTrust() {
            MessageBox.Show("Please enable trust of the Project Object Model", "Project Model Not Trusted",
                    MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
        }

        private void PerformSilently(System.Action action) {
            try {
                Application.Cursor = XlMousePointer.xlWait;
                Application.ScreenUpdating = false;
                Application.DisplayAlerts = false;

                action();
            } finally {
                Application.StatusBar = false;

                Application.DisplayAlerts = true;
                Application.ScreenUpdating = true;
                Application.Cursor = XlMousePointer.xlDefault;
            }
        }

        static Application Application => Globals.ThisAddIn.Application;

        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        private  RibbonGroup        VbASourceExportGroup  { get; }
        private  RibbonToggleButton UseSrcFolderToggle    { get; }
        public   RibbonButton       SelectedProjectButton { get; }
        public   RibbonButton       CurrentProjectButton  { get; }

        IRibbonToggle IVbaSourceExportViewModel.UseSrcFolderToggle    => UseSrcFolderToggle;
        IRibbonButton IVbaSourceExportViewModel.SelectedProjectButton => SelectedProjectButton;
        IRibbonButton IVbaSourceExportViewModel.CurrentProjectButton  => CurrentProjectButton;

        /// <inheritfoc/>
        public Workbook ActiveWorkbook => Application.ActiveWorkbook;

        /// <inheritfoc/>
        public bool     DisplayAlerts { get => Application.DisplayAlerts; set => Application.DisplayAlerts = value; }

        /// <inheritfoc/>
        public dynamic  StatusBar { get => Application.StatusBar; set => Application.StatusBar = value; }

        /// <inheritfoc/>
        public MsoAutomationSecurity AutomationSecurity { get => Application.AutomationSecurity; set => Application.AutomationSecurity = value; }
    }
}

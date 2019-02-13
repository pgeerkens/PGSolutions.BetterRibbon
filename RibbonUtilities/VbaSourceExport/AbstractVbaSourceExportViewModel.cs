////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ComClasses;

namespace PGSolutions.RibbonUtilities.VbaSourceExport {
    using VbaExportSelectedEventHandler = EventHandler<VbaExportSelectedEventArgs>;
    using VbaExportCurrentEventHandler = EventHandler<VbaExportCurrentEventArgs>;

    //using Workbook = Microsoft.Office.Tools.Excel.Workbook;

    /// <summary>.</summary>
    [CLSCompliant(false)]
    public abstract class AbstractVbaSourceExportViewModel : AbstractRibbonGroupViewModel,
            IVbaSourceExportViewModel, IApplication {
        /// <summary>.</summary>
        protected AbstractVbaSourceExportViewModel(IRibbonFactory factory, string suffix) : base(factory) {
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
        public event ToggledEventHandler           UseSrcFolderToggled;
        /// <inheritdoc/>
        public event VbaExportSelectedEventHandler SelectedProjectsClicked;
        /// <inheritdoc/>
        public event VbaExportCurrentEventHandler  CurrentProjectClicked;

        protected virtual void OnSrcFolderToggled(object sender, bool isPressed) {
            UseSrcFolderToggle.SetImageMso(ToggleImage(isPressed));
            UseSrcFolderToggled?.Invoke(sender, isPressed);
        }

        protected virtual void OnExportCurrent(object sender)
        =>  PerformSilently(
                () => CurrentProjectClicked?.Invoke(this,
                        new VbaExportCurrentEventArgs(new ProjectFilterExcel(this), ActiveWorkbook)
            ));

        protected virtual void OnExportSelected(object sender) {
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
                            new VbaExportSelectedEventArgs(list[fd.FilterIndex-1], fd.SelectedItems)
                ));
            }
        }

        protected void PerformSilently(System.Action action) {
            try {
                Application.Cursor = XlMousePointer.xlWait;

                action();
            } finally {
                Application.StatusBar = false;

                Application.Cursor = XlMousePointer.xlDefault;
            }
        }

        private static string ToggleImage(bool isPressed) => isPressed ? "TagMarkComplete" : "MarginsShowHide";

        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        protected RibbonGroup        VbASourceExportGroup  { get; }
        protected RibbonToggleButton UseSrcFolderToggle    { get; }
        /// <inheritdoc/>
        public    RibbonButton       SelectedProjectButton { get; }
        /// <inheritdoc/>
        public    RibbonButton       CurrentProjectButton  { get; }

        IRibbonToggle IVbaSourceExportViewModel.UseSrcFolderToggle    => UseSrcFolderToggle;
        IRibbonButton IVbaSourceExportViewModel.SelectedProjectButton => SelectedProjectButton;
        IRibbonButton IVbaSourceExportViewModel.CurrentProjectButton  => CurrentProjectButton;

        protected abstract Application Application { get; }

        protected abstract Workbook ActiveWorkbook { get; }

        /// <inheritdoc/>
        public virtual void DoOnOpenWorkbook(string wkbkFullName, Action<Workbook> action) {
            if (wkbkFullName == ActiveWorkbook.FullName) {
                action(ActiveWorkbook);
            } else {
                var thisWkbk = ActiveWorkbook;

                Application.DisplayAlerts = false;
                Application.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;

                Application.ScreenUpdating = false;
                var wkbk = Application.Workbooks.Open(wkbkFullName, UpdateLinks:false, ReadOnly:true, AddToMru:false, Editable:false);
                Application.ActiveWindow.Visible = false;
                thisWkbk.Activate();

                try {
                    Application.ScreenUpdating = true;

                    action(wkbk);
                }
                finally {
                    wkbk?.Close(false);

                    Application.DisplayAlerts = true;
                }

            }
        }

        /// <inheritdoc/>
        public abstract bool DisplayAlerts { get; set; }

        /// <inheritdoc/>
        public abstract dynamic StatusBar { get; set; }

        /// <inheritdoc/>
        public abstract MsoAutomationSecurity AutomationSecurity { get; protected set; }
    }
}

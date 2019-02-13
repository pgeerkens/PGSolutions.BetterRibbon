////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ComClasses;

namespace PGSolutions.RibbonUtilities.VbaSourceExport {
    using VbaExportSelectedEventHandler = EventHandler<VbaExportSelectedEventArgs>;
    using VbaExportCurrentEventHandler = EventHandler<VbaExportCurrentEventArgs>;
    using VBE = Microsoft.Vbe.Interop;

    /// <summary>.</summary>
    [CLSCompliant(false)]
    public abstract class AbstractVbaSourceExportViewModel : AbstractRibbonGroupViewModel,
                IVbaSourceExportViewModel, IApplication {
        /// <summary>.</summary>
        protected AbstractVbaSourceExportViewModel(IRibbonFactory factory, string suffix, string itemId, bool isVisible = true, bool isEnabled = true)
        : base(factory, $"VbaExportGroup{suffix}", isVisible, isEnabled) {
            var defaultSize = suffix=="MS" ? false : true;

            UseSrcFolderToggle    = Factory.NewRibbonToggleMso($"UseSrcFolderToggle{suffix}",
                                            isLarge:defaultSize, imageMso:ToggleImage(false));
            SelectedProjectButton = Factory.NewRibbonButtonMso($"SelectedProjectButton{suffix}",
                                            isLarge:defaultSize, imageMso:"SaveAll", showImage:true);
            CurrentProjectButton  = Factory.NewRibbonButtonMso($"CurrentProjectButton{suffix}",
                                            isLarge:defaultSize, imageMso:"FileSaveAs", showImage:true);

            UseSrcFolderToggle.Toggled    += SrcFolderToggled;
            SelectedProjectButton.Attach<RibbonButton>().Clicked += ExportSelected;
            CurrentProjectButton.Attach<RibbonButton>().Clicked  += ExportCurrent;
        }

        /// <inheritdoc/>
        public void Attach(IBooleanSource srcToggleSource) {
            UseSrcFolderToggle.Attach(srcToggleSource.Getter);
            base.Attach();
        }

        /// <inheritdoc/>
        public override void Invalidate() {
            UseSrcFolderToggle.Invalidate();
            base.Invalidate();
        }

        /// <inheritdoc/>
        public event ToggledEventHandler           UseSrcFolderToggled;
        /// <inheritdoc/>
        public event VbaExportSelectedEventHandler SelectedProjectsClicked;
        /// <inheritdoc/>
        public event VbaExportCurrentEventHandler  CurrentProjectClicked;

        protected virtual void SrcFolderToggled(object sender, bool isPressed) {
            UseSrcFolderToggle.SetImageMso(ToggleImage(isPressed));
            UseSrcFolderToggled?.Invoke(sender, isPressed);
        }

        public abstract void ExportCurrent (object sender);
        public abstract void ExportSelected(object sender);

        protected virtual void OnExportCurrent(object sender, VbaExportCurrentEventArgs e)
        => CurrentProjectClicked?.Invoke(this, e);

        protected virtual void OnExportSelected(object sender, VbaExportSelectedEventArgs e)
         => SelectedProjectsClicked?.Invoke(this, e);

        private static string ToggleImage(bool isPressed)
        => isPressed ? "TagMarkComplete" : "MarginsShowHide";

        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
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
        public virtual void DoOnOpenWorkbook(string wkbkFullName, Action<VBE.VBProject, string> action) {
            if (wkbkFullName == ActiveWorkbook.FullName) {
                action?.Invoke(ActiveWorkbook?.VBProject, Path.GetDirectoryName(wkbkFullName));
            } else {
                var thisWkbk = ActiveWorkbook;

                DisplayAlerts = false;
                Application.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;

                Application.ScreenUpdating = false;
                var wkbk = Application.Workbooks.Open(wkbkFullName, UpdateLinks:false, ReadOnly:true,
                        AddToMru:false, Editable:false);
                Application.ActiveWindow.Visible = false;
                thisWkbk.Activate();

                try {
                    Application.ScreenUpdating = true;
                    
                    action?.Invoke(wkbk?.VBProject, Path.GetDirectoryName(wkbkFullName));
                }
                finally {
                    wkbk?.Close(false);

                    DisplayAlerts = true;
                }

            }
        }

        /// <inheritdoc/>
        public abstract bool DisplayAlerts { get; set; }
    }
}

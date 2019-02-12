////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary></summary>
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable")]
    [Description("")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IRibbonGroupModel))]
    [Guid(Guids.RibbonGroupModel)]
    public sealed class RibbonGroupModel :IRibbonGroupModel {
        public RibbonGroupModel(Func<string,RibbonGroup> factory, IRibbonControlStrings strings) {
            Factory = factory;
            Strings = strings;
        }

        public IRibbonGroup ViewModel { get; set; }

        public IRibbonGroupModel Attach(string controlId) {
            var viewModel = Factory(controlId);
            ViewModel = viewModel;
            Invalidate();
            return this;
        }

        private Func<string, RibbonGroup> Factory { get; }

        public IRibbonControlStrings Strings { get; }
        public bool IsEnabled { get; set; } = true;
        public bool IsVisible { get; set; } = true;

        public void Invalidate() {
            if (ViewModel != null) {
                ViewModel.IsEnabled = IsEnabled;
                ViewModel.IsVisible = IsVisible;

                ViewModel.Invalidate();
            }
        }
    }
}

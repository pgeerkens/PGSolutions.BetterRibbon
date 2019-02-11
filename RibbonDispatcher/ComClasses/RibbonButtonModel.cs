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
    [ComSourceInterfaces(typeof(IClickedEvents))]
    [ComDefaultInterface(typeof(IRibbonButtonModel))]
    [Guid(Guids.RibbonButtonModel)]
    public sealed class RibbonButtonModel : IRibbonButtonModel {
        public RibbonButtonModel(Func<string,RibbonButton> factory) => Factory = factory;

        public event ClickedEventHandler Clicked;

        public IRibbonButton ViewModel { get; private set; }

        public void Attach(string controlId, IRibbonControlStrings strings) {
            var viewModel = Factory(controlId);
            viewModel.Attach().SetLanguageStrings(strings);
            viewModel.Clicked += OnClicked;
            ViewModel = viewModel;
        }

        private void OnClicked(object sender) => Clicked(sender);

        private Func<string, RibbonButton> Factory { get; }
    }
}

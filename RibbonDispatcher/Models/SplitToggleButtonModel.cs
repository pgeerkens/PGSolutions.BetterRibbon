////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.Models {
    using IStrings2 = IControlStrings2;

    /// <summary>The COM visible Model for Ribbon Split (Toggle) Button controls.</summary>
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable")]
    [Description("The COM visible Model for Ribbon Split (Toggle) Button controls.")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IToggledEvent))]
    [ComDefaultInterface(typeof(ISplitToggleButtonModel))]
    [Guid(Guids.SplitToggleButtonModel)]
    public class SplitToggleButtonModel: AbstractSplitButtonModel<IToggleSource,ISplitToggleButtonVM>,
            ISplitToggleButtonModel, IToggleSource {
        internal SplitToggleButtonModel(Func<string,SplitToggleButtonVM> funcViewModel,
                IStrings2 strings, ToggleModel toggle, MenuModel menu)
        : base(funcViewModel, strings, menu)
        => _toggleModel = toggle;

        public ISplitToggleButtonModel Attach(string controlId) {
            base.Attach(controlId, this);
            if (ViewModel != null) {
                _toggleModel.Attach(ViewModel.ToggleVM.ControlId);
                _toggleModel.ViewModel.Toggled += OnToggled;
            }
            return this;
        }

        public override void Invalidate() { ToggleModel.Invalidate(); base.Invalidate(); }

        public override void Detach() { ToggleModel.Detach(); base.Detach(); }

        #region Toggleable implementation
        public event ToggledEventHandler Toggled;

        public IToggleModel ToggleModel => _toggleModel; private ToggleModel _toggleModel   { get; }
        public bool        IsPressed { get => ToggleModel.IsPressed; set => ToggleModel.IsPressed = value; }

        private void OnToggled(IRibbonControl control, bool isPressed)
        => Toggled?.Invoke(control, IsPressed = isPressed);
        #endregion

        public ISplitToggleButtonModel SetImage(IImageObject image) {Image = image; return this; }
    }
}

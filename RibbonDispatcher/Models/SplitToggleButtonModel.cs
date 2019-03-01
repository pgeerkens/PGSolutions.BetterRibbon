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
    using IStrings = IControlStrings;

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
                IStrings strings, ToggleModel toggle, MenuModel menu)
        : base(funcViewModel, strings, menu)
        => Toggle = toggle;

        public ISplitToggleButtonModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            if (ViewModel != null) {
                Menu.Attach(ViewModel.MenuVM.Id);

                Toggle.Attach(ViewModel.ToggleVM.Id);
                Toggle.ViewModel.Toggled += OnToggled;
            }
            ViewModel?.Invalidate();
            return this;
        }

        public override void Detach() { Toggle.Detach(); base.Detach(); }

        #region Toggleable implementation
        public event ToggledEventHandler Toggled;

        public ToggleModel Toggle    { get; }
        public bool        IsPressed { get => Toggle.IsPressed; set => Toggle.IsPressed = value; }

        private void OnToggled(IRibbonControl control, bool isPressed)
        => Toggled?.Invoke(control, IsPressed = isPressed);
        #endregion

        public ISplitToggleButtonModel SetImage(ImageObject image) {Image = image; return this; }
    }
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.Models {
    using Microsoft.Office.Core;
    using IStrings = IControlStrings;

    /// <summary>The COM visible Model for Ribbon Button controls.</summary>
    [Description("The COM visible Model for Ribbon Button controls.")]
    [CLSCompliant(true)]
    public abstract class SplitButtonModel<TSource,TControl>: ControlModel<TSource,TControl>,
            ISplitButtonModel
        where TSource: IControlSource where TControl: ISplitButtonVM {
        internal SplitButtonModel(Func<string,IActivatable<TSource,TControl>> funcViewModel, IStrings strings,
                MenuModel menu)
        : base(funcViewModel, strings)
        => Menu   = menu;

        public bool        IsLarge   { get; set; } = true;
        public ImageObject Image     { get; set; } = "MacroSecurity";
        public bool        ShowImage { get; set; } = true;
        public bool        ShowLabel { get; set; } = true;

        public MenuModel   Menu      { get; }
    }

    /// <summary>The COM visible Model for Ribbon Split (Toggle) Button controls.</summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable")]
    [Description("The COM visible Model for Ribbon Split (Toggle) Button controls.")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IToggledEvent))]
    [ComDefaultInterface(typeof(ISplitToggleButtonModel))]
    [Guid(Guids.SplitToggleButtonModel)]
    public class SplitToggleButtonModel: SplitButtonModel<IToggleSource,ISplitToggleButtonVM>,
            ISplitToggleButtonModel, IToggleSource {
        internal SplitToggleButtonModel(Func<string,SplitToggleButtonVM> funcViewModel,
                IStrings strings, ToggleModel toggle, MenuModel menu)
        : base(funcViewModel, strings, menu)
        => Toggle = toggle;

        public ISplitToggleButtonModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            if (ViewModel != null) {
                Toggle.Attach(ViewModel.ToggleVM.Id);
                Menu.Attach(ViewModel.MenuVM.Id);

                Toggle.ViewModel.Toggled += OnToggled;
            }
            ViewModel?.Invalidate();
            return this;
        }

        #region Toggleable implementation
        public event ToggledEventHandler Toggled;

        public ToggleModel Toggle    { get; }
        public bool        IsPressed { get => Toggle.IsPressed; set => Toggle.IsPressed = value; }

        private void OnToggled(IRibbonControl control, bool isPressed)
        => Toggled?.Invoke(control, IsPressed = isPressed);
        #endregion
    }

    /// <summary>The COM visible Model for Ribbon Split (Press) Button controls.</summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable")]
    [Description("The COM visible Model for Ribbon Split (Press) Button controls.")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IClickedEvent))]
    [ComDefaultInterface(typeof(ISplitPressButtonModel))]
    [Guid(Guids.SplitPressButtonModel)]
    public class SplitPressButtonModel: SplitButtonModel<IButtonSource,ISplitPressButtonVM>,
            ISplitPressButtonModel, IButtonSource {
        internal SplitPressButtonModel(Func<string,SplitPressButtonVM> funcViewModel,
                IStrings strings, ButtonModel button, MenuModel menu)
        : base(funcViewModel, strings, menu)
        => Button = button;

        public ISplitPressButtonModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            if (ViewModel != null) {
                Button.Attach(ViewModel.ButtonVM.Id);
                Menu.Attach(ViewModel.MenuVM.Id);

                Button.ViewModel.Clicked += OnClicked;
            }
            ViewModel?.Invalidate();
            return this;
        }

        #region Pressable implementation
        public event ClickedEventHandler Pressed;

        public ButtonModel Button { get; }

        private void OnClicked(IRibbonControl control) => Pressed?.Invoke(control);
        #endregion
    }
}

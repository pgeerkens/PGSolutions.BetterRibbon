////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    ///// <summary></summary>
    //[SuppressMessage( "Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable" )]
    //[Description("")]    
    //[CLSCompliant(true)]
    //[ComVisible(true)]
    //[ClassInterface(ClassInterfaceType.None)]
    //[ComSourceInterfaces(typeof(IToggledEvents))]
    //[ComDefaultInterface(typeof(IRibbonToggleModel))]
    //[Guid(Guids.RibbonToggleModel)]
    //public sealed class RibbonToggleModel : IRibbonToggleModel, IBooleanSource {
    //    internal RibbonToggleModel(ViewModelStore viewModelStore)
    //        => ViewModelStore = viewModelStore;

    //    public event ToggledEventHandler Toggled;

    //    public  IRibbonToggle       ViewModel      => _viewModel;
    //    private RibbonToggleButton  _viewModel;
    //    private ViewModelStore      ViewModelStore { get; }

    //    public bool IsPressed { 
    //        get => _isPressed;
    //        set { _isPressed = value; ViewModel.Invalidate(); }
    //    } private bool _isPressed;

    //    private void OnToggled(object sender, bool isPressed) => Toggled(sender,isPressed);

    //    public void Attach(string controlId, IRibbonControlStrings strings) {
    //        _viewModel = ViewModelStore.AttachToggle(controlId, strings, this);
    //        _viewModel.Toggled += OnToggled;
    //    }

    //    bool IBooleanSource.Getter() => IsPressed;
    //}
}

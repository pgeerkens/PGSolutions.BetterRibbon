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
    //[ComSourceInterfaces(typeof(IClickedEvents))]
    //[ComDefaultInterface(typeof(IRibbonButtonModel))]
    //[Guid(Guids.RibbonButtonModel)]
    //public sealed class RibbonButtonModel : IRibbonButtonModel {
    //    internal RibbonButtonModel(Func<string,IRibbonControlStrings,RibbonButton> factory)
    //        => Factory = factory;

    //    public event ClickedEventHandler Clicked;

    //    public IRibbonButton   ViewModel      => _viewModel;
    //    private RibbonButton   _viewModel     { get; set; }
    //    private ViewModelStore ViewModelStore { get; }
    //    private Func<string,IRibbonControlStrings,RibbonButton> Factory { get; }

    //    public void OnClicked(object sender) => Clicked(sender);

    //    public void Attach(string controlId, IRibbonControlStrings strings) {
    //        _viewModel = Factory(controlId, strings);
    //        _viewModel.Clicked += OnClicked;
    //    }
    //}
}

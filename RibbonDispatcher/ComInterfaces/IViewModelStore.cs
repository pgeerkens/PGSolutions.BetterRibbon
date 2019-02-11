////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary></summary>
    [Description("")]    
    [ComVisible(true)]
    [CLSCompliant(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IViewModelStore)]
    public interface IViewModelStore {
        IRibbonGroup    AttachGroup(string controlId, IRibbonControlStrings strings);
        IRibbonButton   AttachButton(string controlId, IRibbonControlStrings strings);
        IRibbonToggle   AttachToggle(string controlId, IRibbonControlStrings strings, IBooleanSource source);
        IRibbonToggle   AttachCheckBox(string controlId, IRibbonControlStrings strings, IBooleanSource source);
        IRibbonDropDown AttachDropDown(string controlId, IRibbonControlStrings strings, IIntegerSource source);
    }
}

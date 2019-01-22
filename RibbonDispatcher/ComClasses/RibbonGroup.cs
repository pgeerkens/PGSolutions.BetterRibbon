////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary>TODO</summary>
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
       Justification = "Public, Non-Creatable, class.")]
    [Serializable]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IRibbonGroup))]
    [Guid(Guids.RibbonGroup)]
    public class RibbonGroup : RibbonCommon, IRibbonGroup
    {
        internal RibbonGroup(string itemId, IRibbonControlStrings strings, bool visible, bool enabled)
            : base(itemId, strings, visible, enabled) => Attach();
    }
}

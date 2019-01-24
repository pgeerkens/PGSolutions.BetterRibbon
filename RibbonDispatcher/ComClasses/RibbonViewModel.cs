////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary>Implementation of (all) the callbacks for the Fluent Ribbon; for COM clients.</summary>
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
        Justification = "Public Non-Creatable class for COM.")]
    [Serializable]
    [ComVisible(true)]
    [CLSCompliant(true)]
    [ComDefaultInterface(typeof(IRibbonViewModel))]
    [Guid(Guids.RibbonViewModel)]
    public sealed class RibbonViewModel : AbstractRibbonViewModel, IRibbonViewModel {
        /// <summary>TODO</summary>
        internal RibbonViewModel() : base()
            => _id = "TabPGSolutions";

        protected override string Id => _id; private readonly string _id;
    }
}

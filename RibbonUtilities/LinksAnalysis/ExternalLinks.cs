////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

using PGSolutions.RibbonUtilities.LinksAnalysis.Interfaces;

namespace PGSolutions.RibbonUtilities.LinksAnalysis {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IExternalLinks))]
    [Guid(Guids.ExternalLinks2)]
    internal sealed class ExternalLinks :List<ICellRef>, IExternalLinks {
        public ExternalLinks() : base() { }

        int IExternalLinks.Count => Count;

        ICellRef IExternalLinks.Item(int index) => this[index];
    }
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////

using PGSolutions.RibbonDispatcher.ComInterfaces;
using System.Collections.Generic;

namespace PGSolutions.RibbonDispatcher.ComClasses.ViewModels {
    internal interface IContainerControl: IActivatable, IEnumerable<IActivatable> {
        void Add(IActivatable control);
    }
}

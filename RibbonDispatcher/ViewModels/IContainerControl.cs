////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Collections.Generic;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming","CA1710:IdentifiersShouldHaveCorrectSuffix")]
    public interface IContainerControl: IControlVM, IEnumerable<IControlVM> {
        void Add(IControlVM control);

        void PurgeChildren();
    }
}

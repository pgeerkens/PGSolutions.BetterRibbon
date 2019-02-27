////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////

using System.Collections.Generic;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    public interface IContainerControl: IControlVM, IEnumerable<IControlVM> {
        void Add(IControlVM control);

        void PurgeChildren();
    }
}

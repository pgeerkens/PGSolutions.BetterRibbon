////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.Utilities {
    public interface IActivatableControl<TCtl> where TCtl:IRibbonCommon {
        TCtl Attach(IRibbonTextLanguageControl strings);
        void Detach();
    }
}

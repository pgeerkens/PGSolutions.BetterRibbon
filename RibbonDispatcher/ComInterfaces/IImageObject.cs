////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using stdole;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {

    /// <summary></summary>
    public interface IImageObject {
        object Image { get; }
        string ImageMso { get; }
        IPictureDisp ImageDisp { get; }
    }
}

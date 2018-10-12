////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.ExcelRibbon {
    internal class CustomizableButtonsViewModel : AbstractRibbonGroupViewModel {
        public CustomizableButtonsViewModel(IRibbonFactory factory) : base(factory) {
            CustomizableGroup = factory.NewRibbonGroup("CustomizableGroup", true);
            CustomizableButton1 = factory.NewRibbonButtonMso("CustomizableButton1");
            CustomizableButton2 = factory.NewRibbonButtonMso("CustomizableButton2");
            CustomizableButton3 = factory.NewRibbonButtonMso("CustomizableButton3");
        }

        public RibbonGroup  CustomizableGroup   { get; }
        public RibbonButton CustomizableButton1 { get; }
        public RibbonButton CustomizableButton2 { get; }
        public RibbonButton CustomizableButton3 { get; }
    }
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.BetterRibbon {
    internal static partial class RibbonFactoryExtensions {
        public static RibbonGroupViewModel NewCustomButtonsViewModel(this IRibbonFactory factory)
        => factory.NewRibbonGroup("CustomizableGroup")
                .Add<IRibbonToggleSource>(factory.NewRibbonToggle("CustomVbaToggle1"))
                .Add<IRibbonToggleSource>(factory.NewRibbonToggle("CustomVbaToggle2"))
                .Add<IRibbonToggleSource>(factory.NewRibbonToggle("CustomVbaToggle3"))

                .Add<IRibbonToggleSource>(factory.NewRibbonCheckBox("CustomVbaCheckBox1"))
                .Add<IRibbonToggleSource>(factory.NewRibbonCheckBox("CustomVbaCheckBox2"))
                .Add<IRibbonToggleSource>(factory.NewRibbonCheckBox("CustomVbaCheckBox3"))

                .Add<IRibbonDropDownSource>(factory.NewRibbonDropDown("CustomVbaDropDown1"))
                .Add<IRibbonDropDownSource>(factory.NewRibbonDropDown("CustomVbaDropDown2"))
                .Add<IRibbonDropDownSource>(factory.NewRibbonDropDown("CustomVbaDropDown3"))

                .Add<IRibbonButtonSource>(factory.NewRibbonButton("CustomizableButton1"))
                .Add<IRibbonButtonSource>(factory.NewRibbonButton("CustomizableButton2"))
                .Add<IRibbonButtonSource>(factory.NewRibbonButton("CustomizableButton3"));
    }
}

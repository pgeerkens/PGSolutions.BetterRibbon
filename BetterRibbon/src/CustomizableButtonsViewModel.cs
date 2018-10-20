////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Collections.Generic;

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.Utilities;

namespace PGSolutions.BetterRibbon {
    internal class CustomizableButtonsViewModel : AbstractRibbonGroupViewModel {
        public CustomizableButtonsViewModel(IRibbonFactory factory) : base(factory) {
            CustomizableGroup = factory.NewRibbonGroup("CustomizableGroup", true);

            (CustomizableToggle = factory.NewRibbonToggle("CustomVbaToggle")).SetLanguageStrings();

            (CustomizableButton1 = factory.NewRibbonButtonMso("CustomizableButton1")).SetLanguageStrings();
            (CustomizableButton2 = factory.NewRibbonButtonMso("CustomizableButton2")).SetLanguageStrings();
            (CustomizableButton3 = factory.NewRibbonButtonMso("CustomizableButton3")).SetLanguageStrings();

            (CustomizableCheckBox1 = factory.NewRibbonCheckBox("CustomVbaCheckBox1")).SetLanguageStrings();
            (CustomizableCheckBox2 = factory.NewRibbonCheckBox("CustomVbaCheckBox2")).SetLanguageStrings();
            (CustomizableCheckBox3 = factory.NewRibbonCheckBox("CustomVbaCheckBox3")).SetLanguageStrings();

            (CustomizableDropDown1 = factory.NewRibbonDropDown("CustomVbaOptions1")).SetLanguageStrings();
            (CustomizableDropDown2 = factory.NewRibbonDropDown("CustomVbaOptions2")).SetLanguageStrings();
            (CustomizableDropDown3 = factory.NewRibbonDropDown("CustomVbaOptions3")).SetLanguageStrings();

            AdaptorControls = new Dictionary<string, IActivatable>() {
                { CustomizableToggle.Id,    CustomizableToggle },
                { CustomizableButton1.Id,   CustomizableButton1 },
                { CustomizableButton2.Id,   CustomizableButton2 },
                { CustomizableButton3.Id,   CustomizableButton3 },
                { CustomizableCheckBox1.Id, CustomizableCheckBox1 },
                { CustomizableCheckBox2.Id, CustomizableCheckBox2 },
                { CustomizableCheckBox3.Id, CustomizableCheckBox3 },
                { CustomizableDropDown1.Id, CustomizableDropDown1 },
                { CustomizableDropDown2.Id, CustomizableDropDown2 },
                { CustomizableDropDown3.Id, CustomizableDropDown3 }
            };
        }

        public IReadOnlyDictionary<string, IActivatable> AdaptorControls { get; }
        public string             GroupId => CustomizableGroup.Id;

        private RibbonGroup        CustomizableGroup     { get; }
        private RibbonToggleButton CustomizableToggle    { get; }
        private RibbonButton       CustomizableButton1   { get; }
        private RibbonButton       CustomizableButton2   { get; }
        private RibbonButton       CustomizableButton3   { get; }
        private RibbonCheckBox     CustomizableCheckBox1 { get; }
        private RibbonCheckBox     CustomizableCheckBox2 { get; }
        private RibbonCheckBox     CustomizableCheckBox3 { get; }
        private RibbonDropDown     CustomizableDropDown1 { get; }
        private RibbonDropDown     CustomizableDropDown2 { get; }
        private RibbonDropDown     CustomizableDropDown3 { get; }
    }
}

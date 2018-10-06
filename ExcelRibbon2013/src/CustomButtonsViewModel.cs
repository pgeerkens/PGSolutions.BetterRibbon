////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Collections.Generic;

using PGSolutions.RibbonDispatcher;
using PGSolutions.RibbonDispatcher.AbstractCOM;
using PGSolutions.RibbonDispatcher.Concrete;
using static PGSolutions.RibbonDispatcher.AbstractCOM.RdControlSize;

namespace PGSolutions.ExcelRibbon2013 {
    internal class CustomButtonsViewModel : AbstractRibbonGroupViewModel {
        public CustomButtonsViewModel(IRibbonFactory factory) : base(factory) {
            CustomGroup   = factory.NewRibbonGroup("CustomButtonsGroup", false);
            CustomButton1 = factory.NewRibbonButtonMso("AppLaunchButton1", true, true, rdLarge, "RefreshAll", true,  true);
            CustomButton2 = factory.NewRibbonButtonMso("AppLaunchButton2", true, true, rdLarge, "Refresh",    true,  true);
            CustomButton3 = factory.NewRibbonButtonMso("AppLaunchButton3", true, true, rdLarge, "MacroPlay",  true,  true);
            SizeToggle    = factory.NewRibbonToggleMso("SizeToggle",       true, true, rdLarge, NoImage,      false, true);
            ButtonOptions = factory.NewRibbonDropDown("ButtonOptions2",    true, true);
            ButtonOptions.AddItem(factory.NewSelectableItem("LabelOnly"))
                         .AddItem(factory.NewSelectableItem("ImageOnly"))
                         .AddItem(factory.NewSelectableItem("LabelAndImage"));

            CustomButton1.Clicked       += CustomButton1.DefaultButtonAction();
            CustomButton2.Clicked       += CustomButton2.DefaultButtonAction();
            CustomButton3.Clicked       += CustomButton3.DefaultButtonAction();
            SizeToggle.Toggled          += OnToggled;
            ButtonOptions.SelectionMade += OnSelectionMade;

            ButtonOptions.SelectedItemId = "LabelAndImage";
        }

        public RibbonGroup        CustomGroup   { get; }
        public RibbonButton       CustomButton1 { get; }
        public RibbonButton       CustomButton2 { get; }
        public RibbonButton       CustomButton3 { get; }
        public RibbonToggleButton SizeToggle    { get; }
        public RibbonDropDown     ButtonOptions { get; }

        private IList<IRibbonButton> Buttons => new List<IRibbonButton>() { CustomButton1, CustomButton2, CustomButton3 };

        public  void SetVisible(bool isPressed) => CustomGroup.IsVisible = isPressed;

        private void OnToggled(bool isPressed) {
            var _isLarge = !isPressed;
            foreach (var b in Buttons) { b.Size = _isLarge ? rdLarge : rdRegular; }
        }

        private void OnSelectionMade(string selectedId, int selectedIndex) => Buttons.SetView(selectedIndex);
    }
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ControlMixins;
using static PGSolutions.RibbonDispatcher.Utilities.Extensions;

namespace PGSolutions.ExcelRibbon {

    internal class DemonstrationViewModel : AbstractRibbonGroupViewModel, IDemonstrationViewModel {
        public DemonstrationViewModel(IRibbonFactory factory) : base(factory) {
            CustomGroup   = factory.NewRibbonGroup("CustomButtonsGroup", true);
            CustomButton1 = factory.NewRibbonButtonMso("AppLaunchButton1", showImage: true, imageMso:"RefreshAll");
            CustomButton2 = factory.NewRibbonButtonMso("AppLaunchButton2", showImage: true, imageMso:"Refresh");
            CustomButton3 = factory.NewRibbonButtonMso("AppLaunchButton3", showImage: true, imageMso:"MacroPlay");
            IsLargeToggle = factory.NewRibbonToggleMso("SizeToggle",       showImage: true, imageMso:NoImage);
            DisplayOptions = factory.NewRibbonDropDown("ButtonOptions2");
            DisplayOptions.AddItem(factory.NewSelectableItem("LabelOnly"))
                          .AddItem(factory.NewSelectableItem("ImageOnly"))
                          .AddItem(factory.NewSelectableItem("LabelAndImage"));
        }

        public event ToggledEventHandler         IsLargeToggled;
        public event SelectedEventHandler        DisplayOptionSelected;
        public event EventHandler<IRibbonButton> ButtonClicked;

        private RibbonGroup        CustomGroup    { get; }
        private RibbonButton       CustomButton1  { get; }
        private RibbonButton       CustomButton2  { get; }
        private RibbonButton       CustomButton3  { get; }
        private RibbonToggleButton IsLargeToggle  { get; }
        private RibbonDropDown     DisplayOptions { get; }

        private void OnIsLargeToggled(bool isPressed) => IsLargeToggled(isPressed);
        private void OnSelectionMade(string selectedId, int selectedIndex) => DisplayOptionSelected(selectedId, selectedIndex);

        private void OnButton1Clicked() => ButtonClicked(CustomButton1, CustomButton1);
        private void OnButton2Clicked() => ButtonClicked(CustomButton2, CustomButton2);
        private void OnButton3Clicked() => ButtonClicked(CustomButton3, CustomButton3);

        public void Attach(Func<bool> isLargeSource, Func<int> displayOptionSource) {
            DisplayOptions.Attach(displayOptionSource); DisplayOptions.SelectionMade += OnSelectionMade;
            IsLargeToggle.Attach(isLargeSource); IsLargeToggle.Toggled += OnIsLargeToggled;
            CustomButton1.Attach(); CustomButton1.Clicked += OnButton1Clicked;
            CustomButton2.Attach(); CustomButton2.Clicked += OnButton2Clicked;
            CustomButton3.Attach(); CustomButton3.Clicked += OnButton3Clicked;
        }
        public void Detach() {
            CustomButton3.Detach(); CustomButton3.Clicked -= OnButton3Clicked;
            CustomButton2.Detach(); CustomButton2.Clicked -= OnButton2Clicked;
            CustomButton1.Detach(); CustomButton1.Clicked -= OnButton1Clicked;
            IsLargeToggle.Detach(); IsLargeToggle.Toggled -= OnIsLargeToggled;

            DisplayOptions.Detach(); DisplayOptions.SelectionMade -= OnSelectionMade;
        }

        public void Invalidate() {
            DisplayOptions.IsEnabled = ToggleButtonSize(!IsLargeToggle.IsPressed, Buttons);
            Buttons.SetDisplay(DisplayOptions.SelectedItemIndex);
        }

        private IList<IRibbonButton> Buttons => new List<IRibbonButton>() { CustomButton1, CustomButton2, CustomButton3 };
    }
}

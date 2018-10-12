﻿////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Collections.Generic;

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using static PGSolutions.RibbonDispatcher.Extensions;
using static PGSolutions.RibbonDispatcher.ComInterfaces.RdControlSize;

namespace PGSolutions.ExcelRibbon {
    internal class CustomButtonsViewModel : AbstractRibbonGroupViewModel {
        public CustomButtonsViewModel(IRibbonFactory factory) : base(factory) {
            CustomGroup   = factory.NewRibbonGroup("CustomButtonsGroup", true);
            CustomButton1 = factory.NewRibbonButtonMso("AppLaunchButton1", Size:rdLarge, ImageMso:"RefreshAll", ShowImage:true);
            CustomButton2 = factory.NewRibbonButtonMso("AppLaunchButton2", Size:rdLarge, ImageMso:"Refresh",    ShowImage:true);
            CustomButton3 = factory.NewRibbonButtonMso("AppLaunchButton3", Size:rdLarge, ImageMso:"MacroPlay",  ShowImage:true);
            SizeToggle    = factory.NewRibbonToggleMso("SizeToggle",       Size:rdLarge, ImageMso:NoImage,      ShowImage:true);
            ButtonOptions = factory.NewRibbonDropDown("ButtonOptions2");
            ButtonOptions.AddItem(factory.NewSelectableItem("LabelOnly"))
                         .AddItem(factory.NewSelectableItem("ImageOnly"))
                         .AddItem(factory.NewSelectableItem("LabelAndImage"));

            CustomButton1.Clicked       += CustomButton1.DefaultButtonAction();
            CustomButton2.Clicked       += CustomButton2.DefaultButtonAction();
            CustomButton3.Clicked       += CustomButton3.DefaultButtonAction();
            SizeToggle.Toggled          += OnToggled;
            ButtonOptions.SelectionMade += OnSelectionMade;

            ButtonOptions.SelectedItemId = "LabelAndImage";
            ButtonOptions.IsEnabled      = SizeToggle.IsPressed;

            CustomizableGroup   = factory.NewRibbonGroup("CustomizableGroup", true);
            CustomizableButton1 = factory.NewRibbonButtonAdaptorMso("CustomizableButton1", ImageMso:"MacroSecurity");
            CustomizableButton2 = factory.NewRibbonButtonAdaptorMso("CustomizableButton2", ImageMso:"MacroSecurity");
            CustomizableButton3 = factory.NewRibbonButtonAdaptorMso("CustomizableButton3", ImageMso:"MacroSecurity");
        }

        public RibbonGroup        CustomGroup   { get; }
        public RibbonButton       CustomButton1 { get; }
        public RibbonButton       CustomButton2 { get; }
        public RibbonButton       CustomButton3 { get; }
        public RibbonToggleButton SizeToggle    { get; }
        public RibbonDropDown     ButtonOptions { get; }

        public RibbonGroup         CustomizableGroup   { get; }
        public RibbonButtonAdaptor CustomizableButton1 { get; }
        public RibbonButtonAdaptor CustomizableButton2 { get; }
        public RibbonButtonAdaptor CustomizableButton3 { get; }

        private IList<IRibbonButton> Buttons => new List<IRibbonButton>() { CustomButton1, CustomButton2, CustomButton3 };

        private  void SetVisible(bool isVisible) => CustomGroup.IsVisible = isVisible;

        private void OnToggled(bool isPressed) => ButtonOptions.IsEnabled = ToggleButtonSize(!isPressed, Buttons);

        private void OnSelectionMade(string selectedId, int selectedIndex) =>
            Buttons.SetDisplay((LabelImageDisplay)(selectedIndex + 1));
    }
}

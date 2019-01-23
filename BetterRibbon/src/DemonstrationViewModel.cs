////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Collections.Generic;

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using static PGSolutions.RibbonDispatcher.Utilities.Extensions;

namespace PGSolutions.BetterRibbon {

    internal class DemonstrationViewModel : AbstractRibbonGroupViewModel {
        public DemonstrationViewModel(IRibbonFactory factory) : base(factory) {
            CustomGroup    = factory.NewRibbonGroup("CSharpDemoGroup", true);
            CustomButton1  = factory.NewRibbonButtonMso("AppLaunchButton1", showImage:true, imageMso:"RefreshAll");
            CustomButton2  = factory.NewRibbonButtonMso("AppLaunchButton2", showImage:true, imageMso:"Refresh");
            CustomButton3  = factory.NewRibbonButtonMso("AppLaunchButton3", showImage:true, imageMso:"MacroPlay");
            IsLargeToggle  = factory.NewRibbonToggleMso("SizeToggle",       showImage:true, imageMso:NoImage, visible:true, enabled:true);
            DisplayOptions = factory.NewRibbonDropDown("ButtonOptions2");
            DisplayOptions.AddItem(factory.NewSelectableItem("LabelOnly"))
                          .AddItem(factory.NewSelectableItem("ImageOnly"))
                          .AddItem(factory.NewSelectableItem("LabelAndImage"));

            DisplayOptions.SelectionMade += OnDisplaySelection;
            IsLargeToggle.Toggled += OnIsLargeToggled;
            CustomButton1.Clicked += OnButton1Clicked;
            CustomButton2.Clicked += OnButton2Clicked;
            CustomButton3.Clicked += OnButton3Clicked;
        }

        public event ToggledEventHandler  IsLargeToggled;
        public event SelectedEventHandler DisplaySelection;
        public event ClickedEventHandler  Button1Clicked;
        public event ClickedEventHandler  Button2Clicked;
        public event ClickedEventHandler  Button3Clicked;

        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        private RibbonGroup        CustomGroup    { get; }
        private RibbonButton       CustomButton1  { get; }
        private RibbonButton       CustomButton2  { get; }
        private RibbonButton       CustomButton3  { get; }
        private RibbonToggleButton IsLargeToggle  { get; }
        private RibbonDropDown     DisplayOptions { get; }

        public IList<IRibbonButton> Buttons => new List<IRibbonButton>()
                { CustomButton1, CustomButton2, CustomButton3 };

        private void OnIsLargeToggled(object sender, bool ispressed) =>
            IsLargeToggled?.Invoke(sender, ispressed);
        private void OnDisplaySelection(string itemid, int itemindex) =>
            DisplaySelection?.Invoke(itemid, itemindex);

        private void OnButton1Clicked(object sender) => Button1Clicked?.Invoke(sender);
        private void OnButton2Clicked(object sender) => Button2Clicked?.Invoke(sender);
        private void OnButton3Clicked(object sender) => Button3Clicked?.Invoke(sender);

        public void SetButtonSize(bool isLarge) =>
            DisplayOptions.IsEnabled = ! Buttons.SetButtonSize(isLarge);

        public void SetButtonDisplay(LabelImageOptions displayOption) =>
            Buttons.SetDisplay(displayOption);

        public void Attach(Func<bool> isLargeSource, Func<int> selectedItemSource) {
            IsLargeToggle.Attach(isLargeSource);
            DisplayOptions.Attach(selectedItemSource);
            CustomButton1.Attach();
            CustomButton2.Attach();
            CustomButton3.Attach();
        }
        public void Detach() {
            CustomButton3.Detach();
            CustomButton2.Detach();
            CustomButton1.Detach();
            DisplayOptions.Detach();
            IsLargeToggle.Detach();
        }
        public void Invalidate() {
            IsLargeToggle.Invalidate();
            DisplayOptions.Invalidate();
            CustomButton3.Invalidate();
            CustomButton2.Invalidate();
            CustomButton1.Invalidate();
            CustomGroup.Invalidate();
        }
    }
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
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
            CustomGroup    = factory.NewRibbonGroup("CSharpDemoGroup");
            IsLargeToggle  = factory.NewRibbonToggleMso("SizeToggle",       imageMso:NoImage);
            CheckBox1      = factory.NewRibbonCheckBox("CheckBox1", false);
            CheckBox2      = factory.NewRibbonCheckBox("CheckBox2", false);
            CheckBox3      = factory.NewRibbonCheckBox("CheckBox3", false);
            DisplayOptions = factory.NewRibbonDropDown("Dropdown1");
            DisplayOptions.AddItem(factory.NewSelectableItem("LabelOnly"))
                          .AddItem(factory.NewSelectableItem("ImageOnly"))
                          .AddItem(factory.NewSelectableItem("LabelAndImage"));
            Dropdown2      = factory.NewRibbonDropDown("Dropdown2", false);
            Dropdown3      = factory.NewRibbonDropDown("Dropdown3", false);
            CustomButton1  = factory.NewRibbonButtonMso("AppLaunchButton1", imageMso:"RefreshAll");
            CustomButton2  = factory.NewRibbonButtonMso("AppLaunchButton2", imageMso:"Refresh");
            CustomButton3  = factory.NewRibbonButtonMso("AppLaunchButton3", imageMso:"MacroPlay");

            DisplayOptions.SelectionMade += OnDisplaySelection;
            IsLargeToggle.Toggled += OnIsLargeToggled;
            CustomButton1.Clicked += OnButton1Clicked;
            CustomButton2.Clicked += OnButton2Clicked;
            CustomButton3.Clicked += OnButton3Clicked;
        }

        public event EventHandler<bool>  IsLargeToggled;
        public event EventHandler<int>   DisplaySelection;
        public event ClickedEventHandler Button1Clicked;
        public event ClickedEventHandler Button2Clicked;
        public event ClickedEventHandler Button3Clicked;

        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        private RibbonGroup        CustomGroup    { get; }
        private RibbonToggleButton IsLargeToggle  { get; }
        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        private RibbonCheckBox     CheckBox1      { get; }
        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        private RibbonCheckBox     CheckBox2      { get; }
        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        private RibbonCheckBox     CheckBox3      { get; }
        private RibbonDropDown     DisplayOptions { get; }
        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        private RibbonDropDown     Dropdown2      { get; }
        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        private RibbonDropDown     Dropdown3      { get; }
        private RibbonButton       CustomButton1  { get; }
        private RibbonButton       CustomButton2  { get; }
        private RibbonButton       CustomButton3  { get; }

        public IList<IRibbonButton> Buttons => new List<IRibbonButton>()
                { CustomButton1, CustomButton2, CustomButton3 };

        private void OnIsLargeToggled(object sender, bool ispressed) =>
            IsLargeToggled?.Invoke(sender, ispressed);
        private void OnDisplaySelection(object sender, int selectedIndex) =>
            DisplaySelection?.Invoke(sender, selectedIndex);

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
        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
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

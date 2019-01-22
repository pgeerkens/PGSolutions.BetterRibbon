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

            IsLarge        = true;
            DisplayOption  = LabelImageOptions.ShowBoth;

            Attach( () => IsLarge, () => DisplayOption.IndexFromLabelImageDisplay() );
        }

        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        private RibbonGroup        CustomGroup    { get; }
        private RibbonButton       CustomButton1  { get; }
        private RibbonButton       CustomButton2  { get; }
        private RibbonButton       CustomButton3  { get; }
        private RibbonToggleButton IsLargeToggle  { get; }
        private bool               IsLarge        { get; set; }
        private RibbonDropDown     DisplayOptions { get; }
        private LabelImageOptions  DisplayOption  { get; set; }

        private void OnIsLargeToggled(object sender, bool ispressed) {
            IsLarge = !ispressed;
            Invalidate();
        }
        private void OnDisplaySelection(string itemid, int itemindex) {
            DisplayOption = itemindex.IndexToLabelImageDisplay();
            Invalidate();
        }

        private void OnButtonClicked(object sender) => OnButtonClicked(sender as IRibbonButton);
        private void OnButtonClicked(IRibbonButton button) => button?.MsgBoxShow(button?.Id);

        public void Attach(Func<bool> isLargeSource, Func<int> displayOptionSource) {
            DisplayOptions.Attach(displayOptionSource); DisplayOptions.SelectionMade += OnDisplaySelection;
            IsLargeToggle.Attach(isLargeSource); IsLargeToggle.Toggled += OnIsLargeToggled;
            CustomButton1.Attach(); CustomButton1.Clicked += OnButtonClicked;
            CustomButton2.Attach(); CustomButton2.Clicked += OnButtonClicked;
            CustomButton3.Attach(); CustomButton3.Clicked += OnButtonClicked;
        }
        public void Detach() {
            CustomButton3.Detach(); CustomButton3.Clicked -= OnButtonClicked;
            CustomButton2.Detach(); CustomButton2.Clicked -= OnButtonClicked;
            CustomButton1.Detach(); CustomButton1.Clicked -= OnButtonClicked;
            IsLargeToggle.Detach(); IsLargeToggle.Toggled -= OnIsLargeToggled;
            DisplayOptions.Detach(); DisplayOptions.SelectionMade -= OnDisplaySelection;
        }

        public void Invalidate() {
            DisplayOptions.IsEnabled = ToggleButtonSize(!IsLargeToggle.IsPressed, Buttons);
            Buttons.SetDisplay(DisplayOptions.SelectedItemIndex);
        }

        private IList<IRibbonButton> Buttons => new List<IRibbonButton>()
                { CustomButton1, CustomButton2, CustomButton3 };
    }
}

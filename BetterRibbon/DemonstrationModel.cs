////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using PGSolutions.RibbonDispatcher.ComClasses;

using static PGSolutions.RibbonDispatcher.ComClasses.Extensions;

namespace PGSolutions.BetterRibbon {
    internal sealed class DemonstrationModel {
        internal DemonstrationModel(DemonstrationViewModel viewModel) {
            ViewModel      = viewModel;
            IsLarge        = false;
            DisplayOption  = LabelImageOptions.ShowBoth;

         //   ViewModel.Attach();
        }

        private DemonstrationViewModel ViewModel  { get; set; }

        private bool               IsLarge        { get; set; }
        private LabelImageOptions  DisplayOption  { get; set; }

        private void IsLargeToggled(object sender, bool isPressed) {
            IsLarge = isPressed;
            Invalidate();
        }
        private void DisplaySelection(object sender, int itemIndex) {
            DisplayOption = itemIndex.IndexToLabelImageDisplay();
            Invalidate();
        }

        public void Attach() {
            ViewModel.Attach(()=>IsLarge, ()=>DisplayOption.IndexFromLabelImageDisplay());
            ViewModel.IsLargeToggled   += IsLargeToggled;
            ViewModel.DisplaySelection += DisplaySelection;
            ViewModel.Button1Clicked   += DefaultButtonAction;
            ViewModel.Button2Clicked   += DefaultButtonAction;
            ViewModel.Button3Clicked   += DefaultButtonAction;
        }
        public void Detach() {
            ViewModel.Button3Clicked   -= DefaultButtonAction;
            ViewModel.Button2Clicked   -= DefaultButtonAction;
            ViewModel.Button1Clicked   -= DefaultButtonAction;
            ViewModel.DisplaySelection -= DisplaySelection;
            ViewModel.IsLargeToggled   -= IsLargeToggled;
        }
 
        public void Invalidate() {
            ViewModel.SetButtonSize(IsLarge);
            ViewModel.SetButtonDisplay(DisplayOption);
            ViewModel.Invalidate();
        }
    }
}

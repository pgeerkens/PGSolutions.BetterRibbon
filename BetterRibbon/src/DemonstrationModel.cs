////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////

using PGSolutions.RibbonDispatcher.Utilities;
using PGSolutions.RibbonDispatcher.ComInterfaces;

using static PGSolutions.RibbonDispatcher.Utilities.Extensions;

namespace PGSolutions.BetterRibbon {
    internal sealed class DemonstrationModel {
        internal DemonstrationModel(DemonstrationViewModel viewModel) {
            ViewModel      = viewModel;
            IsLarge        = true;
            DisplayOption  = LabelImageOptions.ShowBoth;
        }

        private DemonstrationViewModel ViewModel  { get; set; }

        private bool               IsLarge        { get; set; }
        private LabelImageOptions  DisplayOption  { get; set; }

        private void IsLargeToggled(object sender, bool ispressed) {
            IsLarge = ispressed;
            Invalidate();
        }
        private void DisplaySelection(string itemid, int itemindex) {
            DisplayOption = itemindex.IndexToLabelImageDisplay();
            Invalidate();
        }
        private void ButtonClicked(object sender) => ButtonClicked(sender as IRibbonButton);
        private void ButtonClicked(IRibbonButton button) {
            button?.MsgBoxShow(button?.Id);
            Invalidate();
        }

        public void Attach() {
            ViewModel.Attach(()=>IsLarge, ()=>DisplayOption.IndexFromLabelImageDisplay());
            ViewModel.IsLargeToggled   += IsLargeToggled;
            ViewModel.DisplaySelection += DisplaySelection;
            ViewModel.Button1Clicked   += ButtonClicked;
            ViewModel.Button2Clicked   += ButtonClicked;
            ViewModel.Button3Clicked   += ButtonClicked;
        }
        public void Detach() {
            ViewModel.Button3Clicked   -= ButtonClicked;
            ViewModel.Button2Clicked   -= ButtonClicked;
            ViewModel.Button1Clicked   -= ButtonClicked;
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

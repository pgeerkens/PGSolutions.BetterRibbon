////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////

using System;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ControlMixins;
using static PGSolutions.RibbonDispatcher.Utilities.Extensions;

namespace PGSolutions.ExcelRibbon {
    internal class DemonstrationModel {
        public DemonstrationModel(IDemonstrationViewModel viewmodel) {
            IsRegular     = true;
            DisplayOption = LabelImageDisplay.ShowImage | LabelImageDisplay.ShowLabel;
            ViewModel     = viewmodel;

            ViewModel.IsLargeToggled        += OnIsLargeToggled;
            ViewModel.DisplayOptionSelected += null;
            ViewModel.ButtonClicked         += OnButtonClicked;
            viewmodel.Attach(()=> IsRegular, ()=>DisplayOption);
            ViewModel.Invalidate();
        }
        private bool                    IsRegular     { get; set; }
        private LabelImageDisplay       DisplayOption { get; set; }
        private IDemonstrationViewModel ViewModel     { get; set; }

        private void OnIsLargeToggled(bool isPressed) {
            IsRegular = isPressed;
            ViewModel.Invalidate();
        }
        private void OnDisplaySelection(string itemId, int itemIndex) {
            DisplayOption = (LabelImageDisplay)itemIndex;
            ViewModel.Invalidate();
        }
        private void OnButtonClicked(object sender, IRibbonButton button) => button.MsgBoxShow(button.Id);
    }

    internal interface IDemonstrationViewModel {
        event ToggledEventHandler  IsLargeToggled;
        event SelectedEventHandler DisplayOptionSelected;
        event EventHandler<IRibbonButton> ButtonClicked;

        void Attach(Func<bool> isLargeSource, Func<LabelImageDisplay> displayOption);
        void Detach();

        void Invalidate();
    }
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////

using System;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ComClasses;
using static PGSolutions.RibbonDispatcher.Utilities.Extensions;

namespace PGSolutions.BetterRibbon {
    internal class DemonstrationModel {
        public DemonstrationModel(IDemonstrationViewModel viewmodel) {
            IsRegular     = true;
            DisplayOption = LabelImageOptions.ShowBoth;
            ViewModel     = viewmodel;

            ViewModel.IsLargeToggled        += OnIsLargeToggled;
            ViewModel.DisplayOptionSelected += OnDisplaySelection;
            ViewModel.ButtonClicked         += OnButtonClicked;
            viewmodel.Attach(()=> IsRegular, ()=>DisplayOption.IndexFromLabelImageDisplay());
            ViewModel.Invalidate();
        }
        private bool                    IsRegular     { get; set; }
        private LabelImageOptions       DisplayOption { get; set; }
        private IDemonstrationViewModel ViewModel     { get; set; }

        private void OnIsLargeToggled(object sender, bool isPressed) {
            IsRegular = isPressed;
            ViewModel.Invalidate();
        }
        private void OnDisplaySelection(string itemId, int itemIndex) {
            DisplayOption = itemIndex.IndexToLabelImageDisplay();
            ViewModel.Invalidate();
        }
        private void OnButtonClicked(object sender, IRibbonButton button)
            => button.MsgBoxShow(button.Id);
    }

    internal interface IDemonstrationViewModel {
        event ToggledEventHandler  IsLargeToggled;
        event SelectedEventHandler DisplayOptionSelected;
        event EventHandler<IRibbonButton> ButtonClicked;

        void Attach(Func<bool> isLargeSource, Func<int> displayOption);
        void Detach();

        void Invalidate();
    }
}

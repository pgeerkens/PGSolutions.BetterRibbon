////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using stdole;

namespace PGSolutions.BetterRibbon {
    [CLSCompliant(false)]
    public abstract class AbstractRibbonGroupModel {
        protected AbstractRibbonGroupModel(RibbonGroupViewModel viewModel) {
            ViewModel = viewModel;
        }

        public void Invalidate() => ViewModel.Invalidate();

        protected RibbonGroupViewModel ViewModel { get; }

        protected RibbonButtonModel GetModel<T>(string id, ClickedEventHandler handler, bool isEnabled, bool isVisible, string imageMso)
        where T : RibbonButton {
            var model = new RibbonButtonModel(ViewModel.Add(ViewModel.Factory.NewRibbonButton(id))
                                    .GetControl<T>, GetStrings(id), isEnabled, isVisible);
            model.SetImageMso(imageMso);
            model?.Attach(id);
            model.Clicked += handler;
            return model;
        }

        protected RibbonButtonModel GetModel<T>(string id, ClickedEventHandler handler, bool isEnabled, bool isVisible, IPictureDisp image)
        where T : RibbonButton {
            var model = new RibbonButtonModel(ViewModel.Add(ViewModel.Factory.NewRibbonButton(id))
                                    .GetControl<T>, GetStrings(id), isEnabled, isVisible);
            model.SetImageDisp(image);
            model?.Attach(id);
            model.Clicked += handler;
            return model;
        }

        protected RibbonToggleModel GetModel<T>(string id, ToggledEventHandler handler, bool isEnabled, bool isVisible, string imageMso)
        where T : RibbonCheckBox {
            var model = new RibbonToggleModel(ViewModel.Add(ViewModel.Factory.NewRibbonToggleMso(id, imageMso: imageMso))
                                .GetControl<T>, GetStrings(id), isEnabled, isVisible);
            model.SetImageMso(imageMso);
            model?.Attach(id);
            model.Toggled += handler;
            return model;
        }

    protected RibbonToggleModel GetModel<T>(string id, ToggledEventHandler handler, bool isEnabled, bool isVisible, IPictureDisp image)
        where T : RibbonCheckBox {
            var model = new RibbonToggleModel(ViewModel.Add(ViewModel.Factory.NewRibbonToggle(id, image: image))
                                .GetControl<T>, GetStrings(id), isEnabled, isVisible);
            model.SetImageDisp(image);
            model?.Attach(id);
            model.Toggled += handler;
            return model;
        }

    protected RibbonDropDownModel GetModel<T>(string id, SelectedEventHandler handler, bool isEnabled, bool isVisible)
        where T : RibbonDropDown {
            var model = new RibbonDropDownModel(ViewModel.Add(ViewModel.Factory.NewRibbonDropDown(id))
                                .GetControl<T>, GetStrings(id), isEnabled, isVisible);
            model?.Attach(id);
            model.SelectionMade += handler;
            return model;
        }

        IRibbonControlStrings GetStrings(string id)
        => ViewModel.Factory.ResourceManager.GetControlStrings(id);
    }
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Linq;
using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using stdole;

namespace PGSolutions.BetterRibbon {
    [CLSCompliant(false)]
    public abstract class AbstractRibbonGroupModel2 {
        protected AbstractRibbonGroupModel2(IList<KeyValuePair<string,RibbonGroupViewModel>> viewModels) {
            ViewModels = new List<KeyValuePair<string,RibbonGroupViewModel>>(viewModels);
            ViewModels.ForEach(vm => vm.Value.Attach());
        }

        public void Invalidate() { foreach (var vm in ViewModels) {vm.Value.Invalidate(); } }

        protected List<KeyValuePair<string,RibbonGroupViewModel>> ViewModels { get; }

        protected RibbonButtonModel GetModel<T>(string id, ClickedEventHandler handler, bool isEnabled,
                bool isVisible, string imageMso)
        where T : RibbonButton {
            var model = ( from vm in ViewModels
                          let controlID = $"{id}{vm.Key}"
                          select new RibbonButtonModel(vm.Value.Add(vm.Value.Factory
                                                         .NewRibbonButton(controlID)).GetControl<T>,
                                       GetStrings(vm.Value,controlID), isEnabled, isVisible)
                        ).LastOrDefault();
            if (model != null) {
                model.SetImageMso(imageMso);
                var junk = ( from vm in ViewModels select model.Attach($"{id}{vm.Key}") ).LastOrDefault();
                model.Clicked += handler;
            }
            return model;
        }

        protected RibbonButtonModel GetModel<T>(string id, ClickedEventHandler handler, bool isEnabled,
                bool isVisible, IPictureDisp image)
        where T : RibbonButton {
            var model = ( from vm in ViewModels
                          let controlID = $"{id}{vm.Key}"
                          select new RibbonButtonModel(vm.Value.Add(vm.Value.Factory
                                                         .NewRibbonButton(controlID)).GetControl<T>,
                                       GetStrings(vm.Value,controlID), isEnabled, isVisible)
                        ).LastOrDefault();
            if (model != null) {
                model.SetImageDisp(image);
                var junk = ( from vm in ViewModels select model.Attach($"{id}{vm.Key}") ).LastOrDefault();
                model.Clicked += handler;
            }
            return model;
        }

        protected RibbonToggleModel GetModel<T>(string id, ToggledEventHandler handler, bool isEnabled,
                bool isVisible, string imageMso)
        where T : RibbonCheckBox {
            var model = ( from vm in ViewModels
                          let controlID = $"{id}{vm.Key}"
                          select new RibbonToggleModel(vm.Value.Add(vm.Value.Factory
                                                         .NewRibbonToggle(controlID)).GetControl<T>,
                                       GetStrings(vm.Value,controlID), isEnabled, isVisible)
                        ).LastOrDefault();
            if (model != null) {
                model.SetImageMso(imageMso);
                var junk = ( from vm in ViewModels select model.Attach($"{id}{vm.Key}") ).LastOrDefault();
                model.Toggled += handler;
            }
            return model;
        }

        protected RibbonToggleModel GetModel<T>(string id, ToggledEventHandler handler, bool isEnabled,
                bool isVisible, IPictureDisp image)
        where T : RibbonCheckBox {
            var model = ( from vm in ViewModels
                          let controlID = $"{id}{vm.Key}"
                          select new RibbonToggleModel(vm.Value.Add(vm.Value.Factory
                                                         .NewRibbonToggle(controlID)).GetControl<T>,
                                       GetStrings(vm.Value,controlID), isEnabled, isVisible)
                        ).LastOrDefault();
            if (model != null) {
                model.SetImageDisp(image);
                var junk = ( from vm in ViewModels select model.Attach($"{id}{vm.Key}") ).LastOrDefault();
                model.Toggled += handler;
            }
            return model;
        }

        protected RibbonDropDownModel GetModel<T>(string id, SelectedEventHandler handler, bool isEnabled,
                bool isVisible)
        where T : RibbonDropDown {
            var model = ( from vm in ViewModels
                          let controlID = $"{id}{vm.Key}"
                          select new RibbonDropDownModel(vm.Value.Add(vm.Value.Factory
                                                         .NewRibbonDropDown(controlID)).GetControl<T>,
                                       GetStrings(vm.Value,controlID), isEnabled, isVisible)
                        ).LastOrDefault();
            if (model != null) {
                var junk = ( from vm in ViewModels select model.Attach($"{id}{vm.Key}") ).LastOrDefault();
                model.SelectionMade += handler;
            }
            return model;
        }

        IRibbonControlStrings GetStrings(RibbonGroupViewModel vm, string controlID)
        => vm.Factory.ResourceManager.GetControlStrings(controlID);
    }
}

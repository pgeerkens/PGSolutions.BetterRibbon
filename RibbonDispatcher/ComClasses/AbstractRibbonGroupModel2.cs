////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Collections.Generic;
using System.Linq;
using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using stdole;

namespace PGSolutions.BetterRibbon {
    //[CLSCompliant(false)]
    //public abstract class AbstractRibbonGroupModel2 {
    //    [SuppressMessage("Microsoft.Design", "CA1006:DoNotNestGenericTypesInMemberSignatures")]
    //    protected AbstractRibbonGroupModel2(IList<KeyValuePair<string,RibbonGroupViewModel>> viewModels) {
    //        ViewModels = new List<KeyValuePair<string,RibbonGroupViewModel>>(viewModels);
    //        ViewModels.ForEach(vm => vm.Value.Attach());
    //    }

    //    public void Invalidate() { foreach (var vm in ViewModels) {vm.Value.Invalidate(); } }

    //    [SuppressMessage("Microsoft.Design", "CA1006:DoNotNestGenericTypesInMemberSignatures")]
    //    protected List<KeyValuePair<string,RibbonGroupViewModel>> ViewModels { get; }

    //    [SuppressMessage("Microsoft.Design", "CA1004:GenericMethodsShouldProvideTypeParameter")]
    //    protected RibbonButtonModel GetModel<T>(string id, ClickedEventHandler handler, bool isEnabled,
    //            bool isVisible, string imageMso)
    //    where T : RibbonButton {
    //        var model = ( from vm in ViewModels
    //                      let controlID = $"{id}{vm.Key}"
    //                      select new RibbonButtonModel(vm.Value.Add(vm.Value.Factory
    //                                                     .NewRibbonButton(controlID)).GetControl<T>,
    //                                   GetStrings(vm.Value,controlID), isEnabled, isVisible)
    //                    ).LastOrDefault();
    //        if (model != null) {
    //            model.SetImageMso(imageMso);
    //            ViewModels.ForEach(kvp => model.Attach($"{id}{kvp.Key}"));
    //            model.Clicked += handler;
    //        }
    //        return model;
    //    }

    //    [SuppressMessage("Microsoft.Design", "CA1004:GenericMethodsShouldProvideTypeParameter")]
    //    protected RibbonButtonModel GetModel<T>(string id, ClickedEventHandler handler, bool isEnabled,
    //            bool isVisible, IPictureDisp image)
    //    where T : RibbonButton {
    //        var model = ( from vm in ViewModels
    //                      let controlID = $"{id}{vm.Key}"
    //                      select new RibbonButtonModel(vm.Value.Add(vm.Value.Factory
    //                                                     .NewRibbonButton(controlID)).GetControl<T>,
    //                                   GetStrings(vm.Value,controlID), isEnabled, isVisible)
    //                    ).LastOrDefault();
    //        if (model != null) {
    //            model.SetImageDisp(image);
    //            ViewModels.ForEach(kvp => model.Attach($"{id}{kvp.Key}"));
    //            model.Clicked += handler;
    //        }
    //        return model;
    //    }

    //    [SuppressMessage("Microsoft.Design", "CA1004:GenericMethodsShouldProvideTypeParameter")]
    //    protected RibbonToggleModel GetModel<T>(string id, ToggledEventHandler handler, bool isEnabled,
    //            bool isVisible, string imageMso)
    //    where T : RibbonCheckBox {
    //        var model = ( from vm in ViewModels
    //                      let controlID = $"{id}{vm.Key}"
    //                      select new RibbonToggleModel(vm.Value.Add(vm.Value.Factory
    //                                                     .NewRibbonToggle(controlID)).GetControl<T>,
    //                                   GetStrings(vm.Value,controlID), isEnabled, isVisible)
    //                    ).LastOrDefault();
    //        if (model != null) {
    //            model.SetImageMso(imageMso);
    //            ViewModels.ForEach(kvp => model.Attach($"{id}{kvp.Key}"));
    //            model.Toggled += handler;
    //        }
    //        return model;
    //    }

    //    [SuppressMessage("Microsoft.Design", "CA1004:GenericMethodsShouldProvideTypeParameter")]
    //    protected RibbonToggleModel GetModel<T>(string id, ToggledEventHandler handler, bool isEnabled,
    //            bool isVisible, IPictureDisp image)
    //    where T : RibbonCheckBox {
    //        var model = ( from vm in ViewModels
    //                      let controlID = $"{id}{vm.Key}"
    //                      select new RibbonToggleModel(vm.Value.Add(vm.Value.Factory
    //                                                     .NewRibbonToggle(controlID)).GetControl<T>,
    //                                   GetStrings(vm.Value,controlID), isEnabled, isVisible)
    //                    ).LastOrDefault();
    //        if (model != null) {
    //            model.SetImageDisp(image);
    //            ViewModels.ForEach(kvp => model.Attach($"{id}{kvp.Key}"));
    //            model.Toggled += handler;
    //        }
    //        return model;
    //    }

    //    [SuppressMessage("Microsoft.Design", "CA1004:GenericMethodsShouldProvideTypeParameter")]
    //    protected RibbonDropDownModel GetModel<T>(string id, SelectedEventHandler handler, bool isEnabled,
    //            bool isVisible)
    //    where T : RibbonDropDown {
    //        var model = ( from vm in ViewModels
    //                      let controlID = $"{id}{vm.Key}"
    //                      select new RibbonDropDownModel(vm.Value.Add(vm.Value.Factory
    //                                                     .NewRibbonDropDown(controlID)).GetControl<T>,
    //                                   GetStrings(vm.Value,controlID), isEnabled, isVisible)
    //                    ).LastOrDefault();
    //        if (model != null) {
    //            ViewModels.ForEach(kvp => model.Attach($"{id}{kvp.Key}"));
    //            model.SelectionMade += handler;
    //        }
    //        return model;
    //    }

    //    static IRibbonControlStrings GetStrings(RibbonGroupViewModel vm, string controlID)
    //    => vm.Factory.ResourceManager.GetControlStrings(controlID);
    //}
}

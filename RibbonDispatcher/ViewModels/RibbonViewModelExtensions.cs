////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////

namespace PGSolutions.RibbonDispatcher.ViewModels {
    public static partial class RibbonViewModelExtensions {
        /// <summary>Invalidates the entire Fluent Ribbon.</summary>
        public static void Invalidate(this IRibbonViewModel vm)
        => vm?.RibbonUI?.Invalidate();

        /// <summary>Invalidates this Ribbon Tab.</summary>
        public static void InvalidateTab(this IRibbonViewModel vm)
        => vm?.RibbonUI?.InvalidateControl(vm?.ControlId);

        /// <summary>Invalidates the specified ribbon control.</summary>
        public static void InvalidateControl(this IRibbonViewModel vm, string ControlId)
         => vm?.RibbonUI?.InvalidateControl(ControlId);

        /// <summary>Invalidates the specified Office-Built-In ribbon control.</summary>
        public static void InvalidateControlMso(this IRibbonViewModel vm, string ControlId)
         => vm?.RibbonUI?.InvalidateControlMso(ControlId);

        /// <summary>Activates the specified ribbon tab.</summary>
        public static void ActivateTab(this IRibbonViewModel vm, string ControlId)
         => vm?.RibbonUI?.ActivateTab(ControlId);

        /// <summary>Activates the specified ribbon tab.</summary>
        public static void ActivateTabQ(this IRibbonViewModel vm, string ControlId, string ns)
         => vm?.RibbonUI?.ActivateTabQ(ControlId, ns);
    }
}

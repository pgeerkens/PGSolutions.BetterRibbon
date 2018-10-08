////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Collections.Generic;

using PGSolutions.RibbonDispatcher2013.AbstractCOM;
using PGSolutions.RibbonDispatcher2013.ConcreteCOM;
using PGSolutions.RibbonDispatcher2013.ControlMixins;
using static PGSolutions.RibbonDispatcher2013.AbstractCOM.RdControlSize;

namespace PGSolutions.ExcelRibbon2013 {
    internal class StandardButtonsViewModel : AbstractRibbonGroupViewModel {
        public StandardButtonsViewModel(IRibbonFactory factory, ToggledEventHandler showAdvancedAction) : base(factory) {
            StandardButtonsGroup = Factory.NewRibbonGroup("StandardButtonsGroup");
            StandardButton1      = Factory.NewRibbonButtonMso("StandardButton1",   true, true, rdRegular, "RefreshAll", false, true);
            StandardButton2      = Factory.NewRibbonButtonMso("StandardButton2",   true, true, rdRegular, "Refresh",    false, true);
            ShowAdvancedToggle   = Factory.NewRibbonCheckBox("ShowAdvancedToggle", true, true);
            ButtonOptions        = factory.NewRibbonDropDown("ButtonOptions",      true, true);

            StandardButton1.Clicked     += ExportVba.ExportVbaModules();
            StandardButton2.Clicked     += ExportVba.ExportVbaModulesCurrent();
            ShowAdvancedToggle.Toggled  += showAdvancedAction;
            ButtonOptions.SelectionMade += OnSelectionMade;

            ButtonOptions.OnActionDropDown(null,2);
        }

        public RibbonGroup    StandardButtonsGroup { get; }
        public RibbonButton   StandardButton1      { get; }
        public RibbonButton   StandardButton2      { get; }
        public RibbonDropDown ButtonOptions        { get; }
        public RibbonCheckBox ShowAdvancedToggle   { get; }

        private IList<IRibbonButton> Buttons => new List<IRibbonButton>() { StandardButton1, StandardButton2 };

        private void OnSelectionMade(string selectedId, int selectedIndex) => Buttons.SetView(selectedIndex);
    }
}

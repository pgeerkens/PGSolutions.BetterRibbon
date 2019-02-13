////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher;
using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.BetterRibbon {
    [Serializable]
    [CLSCompliant(false)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ICustomRibbonGroup))]
    [Guid(Guids.RibbonGroup)]
    public class CustomButtonsViewModel : RibbonGroupViewModel, ICustomRibbonGroup {
        public CustomButtonsViewModel(IRibbonFactory factory, bool isVisible = true, bool isEnabled = true)
        : base(factory, "CustomizableGroup", isVisible, isEnabled) {
            (CustomizableToggle1 = factory.NewRibbonToggle("CustomVbaToggle1")).SetLanguageStrings();
            (CustomizableToggle2 = factory.NewRibbonToggle("CustomVbaToggle2")).SetLanguageStrings();
            (CustomizableToggle3 = factory.NewRibbonToggle("CustomVbaToggle3")).SetLanguageStrings();

            (CustomizableCheckBox1 = factory.NewRibbonCheckBox("CustomVbaCheckBox1")).SetLanguageStrings();
            (CustomizableCheckBox2 = factory.NewRibbonCheckBox("CustomVbaCheckBox2")).SetLanguageStrings();
            (CustomizableCheckBox3 = factory.NewRibbonCheckBox("CustomVbaCheckBox3")).SetLanguageStrings();

            (CustomizableDropDown1 = factory.NewRibbonDropDown("CustomVbaDropDown1")).SetLanguageStrings();
            (CustomizableDropDown2 = factory.NewRibbonDropDown("CustomVbaDropDown2")).SetLanguageStrings();
            (CustomizableDropDown3 = factory.NewRibbonDropDown("CustomVbaDropDown3")).SetLanguageStrings();

            (CustomizableButton1 = factory.NewRibbonButtonMso("CustomizableButton1")).SetLanguageStrings();
            (CustomizableButton2 = factory.NewRibbonButtonMso("CustomizableButton2")).SetLanguageStrings();
            (CustomizableButton3 = factory.NewRibbonButtonMso("CustomizableButton3")).SetLanguageStrings();

            AdaptorControls = new Dictionary<string, IActivatable>() {
                { Id,   this },

                { CustomizableToggle1.Id,   CustomizableToggle1 },
                { CustomizableToggle2.Id,   CustomizableToggle2 },
                { CustomizableToggle3.Id,   CustomizableToggle3 },

                { CustomizableCheckBox1.Id, CustomizableCheckBox1 },
                { CustomizableCheckBox2.Id, CustomizableCheckBox2 },
                { CustomizableCheckBox3.Id, CustomizableCheckBox3 },

                { CustomizableDropDown1.Id, CustomizableDropDown1 },
                { CustomizableDropDown2.Id, CustomizableDropDown2 },
                { CustomizableDropDown3.Id, CustomizableDropDown3 },

                { CustomizableButton1.Id,   CustomizableButton1 },
                { CustomizableButton2.Id,   CustomizableButton2 },
                { CustomizableButton3.Id,   CustomizableButton3 }
            };
        }

        public override void Invalidate() {
            CustomizableToggle1?.Invalidate();
            CustomizableToggle2?.Invalidate();
            CustomizableToggle3?.Invalidate();

            CustomizableCheckBox1?.Invalidate();
            CustomizableCheckBox2?.Invalidate();
            CustomizableCheckBox3?.Invalidate();

            CustomizableDropDown1?.Invalidate();
            CustomizableDropDown2?.Invalidate();
            CustomizableDropDown3?.Invalidate();

            CustomizableButton1?.Invalidate();
            CustomizableButton2?.Invalidate();
            CustomizableButton3?.Invalidate();

            base.Invalidate();
        }

        public TControl GetControl<TControl>(string controlId) where TControl:RibbonCommon
        => AdaptorControls.FirstOrDefault(kv => kv.Key == controlId).Value as TControl;

        public void     DetachControls() {
            foreach (var c in AdaptorControls) c.Value.Detach();
        }

        private RibbonToggleButton CustomizableToggle1   { get; }
        private RibbonToggleButton CustomizableToggle2   { get; }
        private RibbonToggleButton CustomizableToggle3   { get; }

        private RibbonCheckBox     CustomizableCheckBox1 { get; }
        private RibbonCheckBox     CustomizableCheckBox2 { get; }
        private RibbonCheckBox     CustomizableCheckBox3 { get; }

        private RibbonDropDown     CustomizableDropDown1 { get; }
        private RibbonDropDown     CustomizableDropDown2 { get; }
        private RibbonDropDown     CustomizableDropDown3 { get; }

        private RibbonButton       CustomizableButton1   { get; }
        private RibbonButton       CustomizableButton2   { get; }
        private RibbonButton       CustomizableButton3   { get; }

        protected IReadOnlyDictionary<string, IActivatable> AdaptorControls { get; }

        /// <inheritdoc/>
        public virtual void SetShowInactive(bool showInactive) {
            foreach (var ctrl in AdaptorControls) {
                ctrl.Value.ShowWhenInactive = showInactive;
                ctrl.Value.Invalidate();
            }
            //Invalidate();
        }
    }
}

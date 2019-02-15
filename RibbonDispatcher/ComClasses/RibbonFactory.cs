////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {

    /// <summary>Implementation of the factory for Ribbon objects.</summary>
    /// <remarks>
    /// The {SuppressMessage} attributes are left in the source here, instead of being 'fired and
    /// forgotten' to the Global Suppresion file, as commentary on a practice often seen as a C#
    /// anti-pattern. Although non-standard C# practice, these "optional parameters with default 
    /// values" usages are (believed to be) the only means of implementing functionality equivalent
    /// to "overrides" in a COM-compatible way.
    /// </remarks>
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
       Justification = "Public, Non-Creatable, class with exported Events.")]
    [Serializable]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IRibbonFactory))]
    [Guid(Guids.RibbonFactory)]
    [Description("Implementation of the factory for Ribbon objects.")]
    public partial class RibbonFactory : IRibbonFactory {
        public RibbonFactory() : this(new ResourceLoader(), null) { ; }

        internal RibbonFactory(IResourceManager manager) : this(null, manager) { ; }

        internal RibbonFactory(ResourceLoader loader, IResourceManager manager) {
            ResourceLoader   = loader;
            ResourceManager  = manager ?? loader;

            _controls    = new Dictionary<string, IRibbonCommon>();
            _sizeables   = new Dictionary<string, ISizeable>();
            _clickables  = new Dictionary<string, IClickable>();
            _toggleables = new Dictionary<string, IToggleable>();
            _selectables = new Dictionary<string, ISelectable>();
            _imageables  = new Dictionary<string, IImageable>();
        }

        internal IResourceLoader  ResourceLoader  { get; }
        /// <inheritdoc/>
        public IResourceManager   ResourceManager { get; }

        private  readonly IDictionary<string, IRibbonCommon> _controls;
        private  readonly IDictionary<string, ISizeable>     _sizeables;
        private  readonly IDictionary<string, IClickable>    _clickables;
        private  readonly IDictionary<string, ISelectable>   _selectables;
        private  readonly IDictionary<string, IImageable>    _imageables;
        private  readonly IDictionary<string, IToggleable>   _toggleables;

        internal object LoadImage(string imageId) => ResourceManager.GetImage(imageId);

        /// <summary>Returns a readonly collection of all Ribbon Controls in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string, IRibbonCommon> Controls    => new ReadOnlyDictionary<string, IRibbonCommon>(_controls);
 
        /// <summary>Returns a readonly collection of all Ribbon (Action) Buttons in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string, ISizeable>     Sizeables   => new ReadOnlyDictionary<string, ISizeable>(_sizeables);

        /// <summary>Returns a readonly collection of all Ribbon (Action) Buttons in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string, IClickable>    Clickables  => new ReadOnlyDictionary<string, IClickable>(_clickables);

        /// <summary>Returns a readonly collection of all Ribbon DropDowns in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string, ISelectable>   Selectables => new ReadOnlyDictionary<string, ISelectable>(_selectables);

        /// <summary>Returns a readonly collection of all Ribbon Imageable Controls in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string, IImageable>    Imageables  => new ReadOnlyDictionary<string, IImageable>(_imageables);

        /// <summary>Returns a readonly collection of all Ribbon Toggle Buttons in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string, IToggleable>   Toggleables => new ReadOnlyDictionary<string, IToggleable>(_toggleables);

        /// <summary>TODO</summary>
        internal event ChangedEventHandler Changed;

        /// <inheritdoc/>
        internal void OnChanged(object sender, IControlChangedEventArgs e) => Changed?.Invoke(this, new ControlChangedEventArgs(e.ControlId));

        public T Add<T,TSource>(T ctrl) where T:RibbonCommon<TSource> where TSource:class,IRibbonCommonSource {
            if (!_controls.ContainsKey(ctrl.Id)) _controls.Add(ctrl.Id, ctrl);

            _clickables .AddNotNull(ctrl.Id, ctrl as IClickable);
            _sizeables  .AddNotNull(ctrl.Id, ctrl as ISizeable);
            _selectables.AddNotNull(ctrl.Id, ctrl as ISelectable);
            _imageables. AddNotNull(ctrl.Id, ctrl as IImageable);
            _toggleables.AddNotNull(ctrl.Id, ctrl as IToggleable);

            ctrl.Changed += OnChanged;
            return ctrl;
        }

        public IRibbonControlStrings GetStrings(string controlId) => ResourceManager.GetControlStrings(controlId);

        /// <summary>Returns a new Ribbon Group ViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public RibbonGroupViewModel NewRibbonGroup(string controlId)
        => Add<RibbonGroupViewModel,IRibbonCommonSource>(new RibbonGroupViewModel(this, controlId));

        /// <summary>Returns a new Ribbon ActionButton ViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public RibbonButton NewRibbonButton(string controlId)
        => Add<RibbonButton,IRibbonButtonSource>(new RibbonButton(controlId));

        /// <summary>Returns a new Ribbon ToggleButton ViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public RibbonToggleButton NewRibbonToggle(string controlId)
        => Add<RibbonToggleButton,IRibbonToggleSource>(new RibbonToggleButton(controlId));

        /// <summary>Returns a new Ribbon CheckBox ViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public RibbonCheckBox NewRibbonCheckBox(string controlId)
        => Add<RibbonCheckBox,IRibbonToggleSource>(new RibbonCheckBox(controlId));

        /// <summary>Returns a new Ribbon DropDownViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public RibbonDropDown NewRibbonDropDown(string controlId)
        => Add<RibbonDropDown,IRibbonDropDownSource>(new RibbonDropDown(controlId));

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public SelectableItem NewSelectableItem(string controlId)
        => new SelectableItem(controlId);

        /// <inheritdoc/>
        public IResourceLoader NewResourceLoader() => ResourceLoader;

        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IRibbonControlStrings NewControlStrings(string label,
                string screenTip = null, string superTip = null, string keyTip = null,
                string alternateLabel = null, string description = null)
        => new RibbonControlStrings(label, screenTip, superTip, keyTip, alternateLabel, description);
    }
}

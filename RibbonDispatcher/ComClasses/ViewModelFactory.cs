////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ComClasses.ViewModels;

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
    [ComDefaultInterface(typeof(IViewModelFactory))]
    [Guid(Guids.RibbonFactory)]
    [Description("Implementation of the factory for Ribbon objects.")]
    public partial class ViewModelFactory : IViewModelFactory {
        public ViewModelFactory() : this(new ResourceLoader()) { ; }

        internal ViewModelFactory(IResourceManager manager) {
            ResourceManager = manager; 

            _controls      = new Dictionary<string, IControlVM>();
            _sizeables     = new Dictionary<string, ISizeableVM>();
            _clickables    = new Dictionary<string, IClickableVM>();
            _selectables   = new Dictionary<string, ISelectableVM>();
            _imageables    = new Dictionary<string, IImageableVM>();
            _toggleables   = new Dictionary<string, IToggleableVM>();
            _textEditables = new Dictionary<string, IEditableVM>();
        }

        /// <inheritdoc/>
        public IResourceManager   ResourceManager { get; }

        private  readonly IDictionary<string, IControlVM>    _controls;
        private  readonly IDictionary<string, ISizeableVM>   _sizeables;
        private  readonly IDictionary<string, IClickableVM>  _clickables;
        private  readonly IDictionary<string, ISelectableVM> _selectables;
        private  readonly IDictionary<string, IImageableVM>  _imageables;
        private  readonly IDictionary<string, IToggleableVM> _toggleables;
        private  readonly IDictionary<string, IEditableVM>   _textEditables;

        public object LoadImage(string imageId) => ResourceManager.GetImage(imageId);

        /// <summary>Returns a readonly collection of all Ribbon Controls in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string, IControlVM>    Controls      => new ReadOnlyDictionary<string, IControlVM>(_controls);
 
        /// <summary>Returns a readonly collection of all Ribbon (Action) Buttons in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string, ISizeableVM>   Sizeables     => new ReadOnlyDictionary<string, ISizeableVM>(_sizeables);

        /// <summary>Returns a readonly collection of all Ribbon (Action) Buttons in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string, IClickableVM>  Clickables    => new ReadOnlyDictionary<string, IClickableVM>(_clickables);

        /// <summary>Returns a readonly collection of all Ribbon DropDowns in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string, ISelectableVM> Selectables   => new ReadOnlyDictionary<string, ISelectableVM>(_selectables);

        /// <summary>Returns a readonly collection of all Ribbon Imageable Controls in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string, IImageableVM>  Imageables    => new ReadOnlyDictionary<string, IImageableVM>(_imageables);

        /// <summary>Returns a readonly collection of all Ribbon Toggle Buttons in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string, IToggleableVM> Toggleables   => new ReadOnlyDictionary<string, IToggleableVM>(_toggleables);

        /// <summary>Returns a readonly collection of all Ribbon Toggle Buttons in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string, IEditableVM>   TextEditables => new ReadOnlyDictionary<string, IEditableVM>(_textEditables);

        /// <inheritdoc/>
        internal TControl GetControl<TControl>(string controlId) where TControl : class, IControlVM
        => Controls.FirstOrDefault( c => c.Key == controlId).Value as TControl;

        /// <inheritdoc/>
        internal event ChangedEventHandler Changed;

        /// <inheritdoc/>
        internal void OnChanged(object sender, IControlChangedEventArgs e) => Changed?.Invoke(this, new ControlChangedEventArgs(e.ControlId));

        [SuppressMessage("Microsoft.Design", "CA1004:GenericMethodsShouldProvideTypeParameter")]
        internal T Add<T,TSource>(T ctrl) where T:AbstractControlVM<TSource> where TSource:class,IRibbonCommonSource {
            if (!_controls.ContainsKey(ctrl.Id)) _controls.Add(ctrl.Id, ctrl);

            _clickables   .AddNotNull(ctrl.Id, ctrl as IClickableVM);
            _sizeables    .AddNotNull(ctrl.Id, ctrl as ISizeableVM);
            _selectables  .AddNotNull(ctrl.Id, ctrl as ISelectableVM);
            _imageables   .AddNotNull(ctrl.Id, ctrl as IImageableVM);
            _toggleables  .AddNotNull(ctrl.Id, ctrl as IToggleableVM);
            _textEditables.AddNotNull(ctrl.Id, ctrl as IEditableVM);

            ctrl.Changed += OnChanged;
            return ctrl;
        }

        public IControlStrings GetStrings(string controlId) => ResourceManager.GetControlStrings(controlId);

        /// <summary>Returns a new Ribbon Group ViewModel instance.</summary>
        internal GroupVM NewGroup(string controlId)
        => Add<GroupVM,IRibbonCommonSource>(new GroupVM(this, controlId));

        /// <summary>Returns a new Ribbon ActionButton ViewModel instance.</summary>
        internal ButtonVM NewButton(string controlId)
        => Add<ButtonVM,IButtonSource>(new ButtonVM(controlId));

        /// <summary>Returns a new Ribbon ToggleButton ViewModel instance.</summary>
        internal ToggleButtonVM NewToggleButton(string controlId)
        => Add<ToggleButtonVM,IToggleSource>(new ToggleButtonVM(controlId));

        /// <summary>Returns a new Ribbon CheckBoxVM ViewModel instance.</summary>
        internal CheckBoxVM NewCheckBox(string controlId)
        => Add<CheckBoxVM,IToggleSource>(new CheckBoxVM(controlId));

        /// <summary>Returns a new Ribbon DropDownViewModel instance.</summary>
        internal DropDownVM NewDropDown(string controlId)
        => Add<DropDownVM,IDropDownSource>(new DropDownVM(controlId));

        /// <inheritdoc/>
        internal SelectableItem NewSelectableItem(string controlId)
        => new SelectableItem(controlId);

        /// <summary>Returns a new Ribbon ToggleButton ViewModel instance.</summary>
        internal EditBoxVM NewEditBox(string controlId)
        => Add<EditBoxVM, IEditBoxSource>(new EditBoxVM(controlId));

        /// <summary>Returns a new Ribbon ToggleButton ViewModel instance.</summary>
        internal ComboBoxVM NewComboBox(string controlId)
        => Add<ComboBoxVM, IComboBoxSource>(new ComboBoxVM(controlId));

        ///// <inheritdoc/>
        //public IResourceLoader NewResourceLoader() => ResourceLoader;

        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IControlStrings NewControlStrings(string label,
                string screenTip = null, string superTip = null, string keyTip = null,
                string alternateLabel = null, string description = null)
        => new ControlStrings(label, screenTip, superTip, keyTip, alternateLabel, description);
    }
}

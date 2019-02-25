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
    /// 
    /// This class must be COM-Visible for the typelib to be created. 
    /// </remarks>
    [Description("Implementation of the factory for Ribbon objects.")]
    [Serializable]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IViewModelFactory))]
    [Guid(Guids.ViewModelFactory)]
    public partial class ViewModelFactory : IViewModelFactory {
        public ViewModelFactory() : this(new ResourceLoader()) { }

        internal ViewModelFactory(IResourceManager manager) {
            ResourceManager = manager; 

            _controls      = new Dictionary<string, IControlVM>();
            _sizeables     = new Dictionary<string, ISizeableVM>();
            _clickables    = new Dictionary<string, IClickableVM>();
            _selectables   = new Dictionary<string, ISelectableVM>();
            _selectables2  = new Dictionary<string, ISelectable2VM>();
            _imageables    = new Dictionary<string, IImageableVM>();
            _toggleables   = new Dictionary<string, IToggleableVM>();
            _textEditables = new Dictionary<string, IEditableVM>();
            _dynamicMenus  = new Dictionary<string, IDynamicMenuVM>();
            _gallerySizes  = new Dictionary<string, IGallerySizeVM>();
        }

        /// <inheritdoc/>
        internal event ChangedEventHandler Changed;

        /// <inheritdoc/>
        internal void OnChanged(object sender, IControlChangedEventArgs e) => Changed?.Invoke(this, new ControlChangedEventArgs(e.ControlId));

        /// <inheritdoc/>
        public IResourceManager ResourceManager { get; }

        /// <inheritdoc/>
        internal TControl GetControl<TControl>(string controlId) where TControl : class, IControlVM
        => Controls.FirstOrDefault(c => c.Key == controlId).Value as TControl;

        #region IVewModelFactory implementation
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IControlStrings NewControlStrings(string label,
                                                 string screenTip      = null,
                                                 string superTip       = null,
                                                 string keyTip         = null,
                                                 string alternateLabel = null,
                                                 string description    = null)
        => new ControlStrings(label, screenTip, superTip, keyTip, alternateLabel, description);

        public IControlStrings GetStrings(string controlId) => ResourceManager.GetControlStrings(controlId);

        public object LoadImage(string imageId) => ResourceManager.GetImage(imageId);
        #endregion

        #region Dictionaries
        private  readonly IDictionary<string, IControlVM>     _controls;
        private  readonly IDictionary<string, ISizeableVM>    _sizeables;
        private  readonly IDictionary<string, IClickableVM>   _clickables;
        private  readonly IDictionary<string, ISelectableVM>  _selectables;
        private  readonly IDictionary<string, ISelectable2VM> _selectables2;
        private  readonly IDictionary<string, IImageableVM>   _imageables;
        private  readonly IDictionary<string, IToggleableVM>  _toggleables;
        private  readonly IDictionary<string, IEditableVM>    _textEditables;
        private  readonly IDictionary<string, IDynamicMenuVM> _dynamicMenus;
        private  readonly IDictionary<string, IGallerySizeVM> _gallerySizes;

        /// <summary>Returns a readonly collection of all Ribbon Controls in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string, IControlVM>    Controls      => new ReadOnlyDictionary<string, IControlVM>(_controls);
 
        /// <summary>Returns a readonly collection of all Ribbon (Action) Buttons in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string, ISizeableVM>   Sizeables     => new ReadOnlyDictionary<string, ISizeableVM>(_sizeables);

        /// <summary>Returns a readonly collection of all Ribbon (Action) Buttons in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string, IClickableVM>  Clickables    => new ReadOnlyDictionary<string, IClickableVM>(_clickables);

        /// <summary>Returns a readonly collection of all Ribbon DropDowns in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string, ISelectableVM> Selectables   => new ReadOnlyDictionary<string, ISelectableVM>(_selectables);

        /// <summary>Returns a readonly collection of all Ribbon DropDowns in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string, ISelectable2VM> Selectables2 => new ReadOnlyDictionary<string, ISelectable2VM>(_selectables2);

        /// <summary>Returns a readonly collection of all Ribbon Imageable Controls in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string, IImageableVM>  Imageables    => new ReadOnlyDictionary<string, IImageableVM>(_imageables);

        /// <summary>Returns a readonly collection of all Ribbon Toggle Buttons in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string, IToggleableVM> Toggleables   => new ReadOnlyDictionary<string, IToggleableVM>(_toggleables);

        /// <summary>Returns a readonly collection of all Ribbon Toggle Buttons in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string, IEditableVM>   TextEditables => new ReadOnlyDictionary<string, IEditableVM>(_textEditables);

        /// <summary>Returns a readonly collection of all Ribbon Toggle Buttons in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string, IDynamicMenuVM> DynamicMenus => new ReadOnlyDictionary<string, IDynamicMenuVM>(_dynamicMenus);

        /// <summary>Returns a readonly collection of all Ribbon Toggle Buttons in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string, IGallerySizeVM> GallerySizes => new ReadOnlyDictionary<string, IGallerySizeVM>(_gallerySizes);
        #endregion

        #region Factoy Method implementation
        [SuppressMessage("Microsoft.Design", "CA1004:GenericMethodsShouldProvideTypeParameter")]
        internal T Add<T, TSource>(T ctrl) where T : AbstractControlVM<TSource> where TSource : class, IControlSource {
            if (!_controls.ContainsKey(ctrl.Id)) _controls.Add(ctrl.Id, ctrl);

            _clickables.AddNotNull(ctrl.Id, ctrl as IClickableVM);
            _sizeables.AddNotNull(ctrl.Id, ctrl as ISizeableVM);
            _selectables.AddNotNull(ctrl.Id, ctrl as ISelectableVM);
            _selectables2.AddNotNull(ctrl.Id, ctrl as ISelectable2VM);
            _imageables.AddNotNull(ctrl.Id, ctrl as IImageableVM);
            _toggleables.AddNotNull(ctrl.Id, ctrl as IToggleableVM);
            _textEditables.AddNotNull(ctrl.Id, ctrl as IEditableVM);

            ctrl.Changed += OnChanged;
            return ctrl;
        }

        /// <summary>Returns a new Ribbon Group ViewModel instance.</summary>
        internal TabVM NewTab(string controlId)
        => Add<TabVM, IControlSource>(new TabVM(this, controlId));

        /// <summary>Returns a new Ribbon Group ViewModel instance.</summary>
        internal GroupVM NewGroup(string controlId)
        => Add<GroupVM,IControlSource>(new GroupVM(this, controlId));

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
        internal SelectableItemVM NewSelectableItem(string controlId)
        => new SelectableItemVM(controlId);

        /// <summary>Returns a new Ribbon ToggleButton ViewModel instance.</summary>
        internal EditBoxVM NewEditBox(string controlId)
        => Add<EditBoxVM, IEditBoxSource>(new EditBoxVM(controlId));

        /// <summary>Returns a new Ribbon ToggleButton ViewModel instance.</summary>
        internal ComboBoxVM NewComboBox(string controlId)
        => Add<ComboBoxVM, IComboBoxSource>(new ComboBoxVM(controlId));

        /// <summary>Returns a new Ribbon ToggleButton ViewModel instance.</summary>
        internal LabelVM NewLabel(string controlId)
        => Add<LabelVM, ILabelSource>(new LabelVM(controlId));

        /// <summary>Returns a new Ribbon ToggleButton ViewModel instance.</summary>
        internal SplitButtonVM NewSplitButton(string controlId, IButtonVM button, IMenuVM menu)
        => Add<SplitButtonVM, ISplitButtonSource>(new SplitButtonVM(this, controlId, button, menu));

        /// <summary>Returns a new Ribbon ToggleButton ViewModel instance.</summary>
        internal MenuVM NewMenu(string controlId)
        => Add<MenuVM, IMenuSource>(new MenuVM(this, controlId));
        #endregion
    }
}

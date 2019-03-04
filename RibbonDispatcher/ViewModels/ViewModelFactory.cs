////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    /// <summary>Implementation of the factory for Ribbon objects.</summary>
    /// <remarks>
    /// 
    /// This class must be COM-Visible for the typelib to be created. 
    /// 
    /// </remarks>
    [CLSCompliant(true)]
    [Description("The view-model factory for Ribbon objects. Visible to enable TypeLib creation.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IViewModelFactory))]
    [ComVisible(true), Serializable, Guid(Guids.ViewModelFactory)]
    public partial class ViewModelFactory : IViewModelFactory {
        // The nature of this class and constructor ensures automated TypeLib creation.
        public ViewModelFactory() {
            _controls        = new Dictionary<string,IControlVM>();
            _editables       = new Dictionary<string,IEditableVM>();
            _sizeables       = new Dictionary<string,ISizeableVM>();
            _imageables      = new Dictionary<string,IImageableVM>();
            _clickables      = new Dictionary<string,IClickableVM>();
            _toggleables     = new Dictionary<string,IToggleableVM>();
            _selectItems     = new Dictionary<string,ISelectItemsVM>();
            _selectables     = new Dictionary<string,ISelectablesVM>();
            _dynamicMenus    = new Dictionary<string,IDynamicMenuVM>();
            _gallerySizes    = new Dictionary<string,IGallerySizeVM>();
            _menuSeparators  = new Dictionary<string,IMenuSeparatorVM>();
            _descriptionable = new Dictionary<string,IDescriptionableVM>();

            TabViewModels    = new KeyedControls();
        }

        /// <summary>.</summary>
        internal event ChangedEventHandler Changed;

        public void ClearChangedListeners() => Changed = null;

        /// <summary>.</summary>
        internal void OnChanged(object sender, ControlChangedEventArgs e)
        => Changed?.Invoke(this, e);

        /// <summary> TODO Is this needed? Where used, really? </summary>
        public KeyedControls TabViewModels { get; }

        /// <summary>.</summary>
        public TControl GetControl<TControl>(string controlId) where TControl : class, IControlVM
        => Controls[controlId] as TControl;

        #region Dictionaries
        /// <summary>Returns a readonly collection of all Ribbon Controls in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string,IControlVM>         Controls         => new ReadOnlyDictionary<string,IControlVM>(_controls);

        /// <summary>Returns a readonly collection of all Ribbon Toggle Buttons in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string,IEditableVM>        Editables        => new ReadOnlyDictionary<string,IEditableVM>(_editables);

        /// <summary>Returns a readonly collection of all Ribbon (Action) Buttons in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string,ISizeableVM>        Sizeables        => new ReadOnlyDictionary<string,ISizeableVM>(_sizeables);

        /// <summary>Returns a readonly collection of all Ribbon Imageable Controls in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string,IImageableVM>       Imageables       => new ReadOnlyDictionary<string,IImageableVM>(_imageables);

        /// <summary>Returns a readonly collection of all Ribbon (Action) Buttons in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string,IClickableVM>       Clickables       => new ReadOnlyDictionary<string,IClickableVM>(_clickables);

        /// <summary>Returns a readonly collection of all Ribbon Toggle Buttons in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string,IToggleableVM>      Toggleables      => new ReadOnlyDictionary<string,IToggleableVM>(_toggleables);

        /// <summary>Returns a readonly collection of all Ribbon DropDowns in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string,ISelectItemsVM>      SelectItems      => new ReadOnlyDictionary<string,ISelectItemsVM>(_selectItems);

        /// <summary>Returns a readonly collection of all Ribbon DropDowns in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string,ISelectablesVM>     Selectable2      => new ReadOnlyDictionary<string,ISelectablesVM>(_selectables);

        /// <summary>Returns a readonly collection of all Ribbon Toggle Buttons in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string,IDynamicMenuVM>     DynamicMenus     => new ReadOnlyDictionary<string,IDynamicMenuVM>(_dynamicMenus);

        /// <summary>Returns a readonly collection of all Ribbon Toggle Buttons in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string,IGallerySizeVM>     GallerySizes     => new ReadOnlyDictionary<string,IGallerySizeVM>(_gallerySizes);

        /// <summary>Returns a readonly collection of all Ribbon Toggle Buttons in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string,IMenuSeparatorVM>   MenuSeparators   => new ReadOnlyDictionary<string,IMenuSeparatorVM>(_menuSeparators);

        /// <summary>Returns a readonly collection of all Ribbon Toggle Buttons in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string,IDescriptionableVM> Descriptionables => new ReadOnlyDictionary<string,IDescriptionableVM>(_descriptionable);

        private  readonly IDictionary<string, IControlVM>         _controls;
        private  readonly IDictionary<string, IEditableVM>        _editables;
        private  readonly IDictionary<string, ISizeableVM>        _sizeables;
        private  readonly IDictionary<string, IImageableVM>       _imageables;
        private  readonly IDictionary<string, IClickableVM>       _clickables;
        private  readonly IDictionary<string, IToggleableVM>      _toggleables;
        private  readonly IDictionary<string, ISelectItemsVM>     _selectItems;
        private  readonly IDictionary<string, ISelectablesVM>     _selectables;
        private  readonly IDictionary<string, IDynamicMenuVM>     _dynamicMenus;
        private  readonly IDictionary<string, IGallerySizeVM>     _gallerySizes;
        private  readonly IDictionary<string, IMenuSeparatorVM>   _menuSeparators;
        private  readonly IDictionary<string, IDescriptionableVM> _descriptionable;

        [SuppressMessage("Microsoft.Design", "CA1004:GenericMethodsShouldProvideTypeParameter")]
        private TControl Add<TControl,TSource,TVM>(TControl ctrl) 
        where TControl: AbstractControlVM<TSource,TVM>
        where TSource: class,IControlSource
        where TVM:class,IControlVM {
            if (!_controls.ContainsKey(ctrl.ControlId)) {
                _controls       .Add(ctrl.ControlId, ctrl);
                _editables      .AddNotNull(ctrl.ControlId, ctrl as IEditableVM);
                _sizeables      .AddNotNull(ctrl.ControlId, ctrl as ISizeableVM);
                _imageables     .AddNotNull(ctrl.ControlId, ctrl as IImageableVM);
                _clickables     .AddNotNull(ctrl.ControlId, ctrl as IClickableVM);
                _toggleables    .AddNotNull(ctrl.ControlId, ctrl as IToggleableVM);
                _selectItems    .AddNotNull(ctrl.ControlId, ctrl as ISelectItemsVM);
                _selectables    .AddNotNull(ctrl.ControlId, ctrl as ISelectablesVM);
                _dynamicMenus   .AddNotNull(ctrl.ControlId, ctrl as IDynamicMenuVM);
                _gallerySizes   .AddNotNull(ctrl.ControlId, ctrl as IGallerySizeVM);
                _menuSeparators .AddNotNull(ctrl.ControlId, ctrl as IMenuSeparatorVM);
                _descriptionable.AddNotNull(ctrl.ControlId, ctrl as IDescriptionableVM);

                if (ctrl is ITabVM tab) TabViewModels.Add(tab); // TODO - Is this needed?
                ctrl.Changed += OnChanged;
            }
            return ctrl;
        }

        internal void Remove(IContainerControl requestor, IControlVM ctrl) {
            if (requestor==null  ||  ctrl == null) return;

            if (_descriptionable.ContainsKey(ctrl.ControlId)) _descriptionable.Remove(ctrl.ControlId);
            if (_menuSeparators .ContainsKey(ctrl.ControlId)) _menuSeparators .Remove(ctrl.ControlId);
            if (_gallerySizes   .ContainsKey(ctrl.ControlId)) _gallerySizes   .Remove(ctrl.ControlId);
            if (_dynamicMenus   .ContainsKey(ctrl.ControlId)) _dynamicMenus   .Remove(ctrl.ControlId);
            if (_selectables    .ContainsKey(ctrl.ControlId)) _selectables    .Remove(ctrl.ControlId);
            if (_selectItems    .ContainsKey(ctrl.ControlId)) _selectItems    .Remove(ctrl.ControlId);
            if (_toggleables    .ContainsKey(ctrl.ControlId)) _toggleables    .Remove(ctrl.ControlId);
            if (_clickables     .ContainsKey(ctrl.ControlId)) _clickables     .Remove(ctrl.ControlId);
            if (_imageables     .ContainsKey(ctrl.ControlId)) _imageables     .Remove(ctrl.ControlId);
            if (_sizeables      .ContainsKey(ctrl.ControlId)) _sizeables      .Remove(ctrl.ControlId);
            if (_editables      .ContainsKey(ctrl.ControlId)) _editables      .Remove(ctrl.ControlId);
            if (_controls       .ContainsKey(ctrl.ControlId)) _controls       .Remove(ctrl.ControlId);

            ctrl.OnPurged(requestor);
        }
        #endregion

        #region Factoy Method implementation
        /// <summary>.</summary>
        internal TabVM NewTab(string controlId)
        => Add<TabVM,IControlSource,ITabVM>(new TabVM(this, controlId));

        /// <summary>Returns a new Ribbon Group view-model instance.</summary>
        internal GroupVM NewGroup(string controlId, IEnumerable<IControlVM> controls)
        => Add<GroupVM,IControlSource,IGroupVM>(new GroupVM(controlId, controls));

        /// <summary>Returns a new Ribbon ActionButton view-model instance.</summary>
        internal ButtonVM NewButton(string controlId)
        => Add<ButtonVM,IButtonSource,IButtonVM>(new ButtonVM(controlId));

        /// <summary>Returns a new Ribbon ToggleButton view-model instance.</summary>
        internal ToggleButtonVM NewToggleButton(string controlId)
        => Add<ToggleButtonVM,IToggleSource,IToggleVM>(new ToggleButtonVM(controlId));

        /// <summary>Returns a new Ribbon CheckBox view-model instance.</summary>
        internal CheckBoxVM NewCheckBox(string controlId)
        => Add<CheckBoxVM,IToggleSource,IToggleVM>(new CheckBoxVM(controlId));

        /// <summary>Returns a new Ribbon DropDown view-model instance.</summary>
        internal IDropDownVM NewDropDown(string controlId, IReadOnlyList<StaticItemVM> items)
        => (items?.Count ?? 0) > 0
            ? Add<StaticDropDownVM,IStaticDropDownSource,IDropDownVM>(new StaticDropDownVM(controlId,items))
            : Add<DropDownVM,IDropDownSource,IDropDownVM>(new DropDownVM(controlId)) as IDropDownVM;

        /// <summary>Returns a new Ribbon SelectableItem view-model instance.</summary>
        [SuppressMessage("Microsoft.Performance","CA1822:MarkMembersAsStatic")]
        internal StaticItemVM NewStaticItem(string controlId, IControlStrings strings)
        => new StaticItemVM(controlId, strings);

        /// <summary>Returns a new Ribbon EditBox view-model instance.</summary>
        internal EditBoxVM NewEditBox(string controlId)
        => Add<EditBoxVM, IEditBoxSource,IEditBoxVM>(new EditBoxVM(controlId));

        /// <summary>Returns a new Ribbon ComboBox view-model instance.</summary>
        internal IControlVM NewComboBox(string controlId, IReadOnlyList<StaticItemVM> items)
        => (items?.Count ?? 0) > 0
            ? Add<StaticComboBoxVM, IStaticComboBoxSource,IStaticComboBoxVM>(new StaticComboBoxVM(controlId,items))
            : Add<ComboBoxVM, IComboBoxSource,IComboBoxVM>(new ComboBoxVM(controlId)) as IControlVM;

        /// <summary>Returns a new Ribbon ComboBox view-model instance.</summary>
        internal IControlVM NewGallery(string controlId, IReadOnlyList<StaticItemVM> items)
        => (items?.Count ?? 0) > 0
            ? Add<StaticGalleryVM, IStaticGallerySource,IStaticGalleryVM>(new StaticGalleryVM(controlId,items))
            : Add<GalleryVM, IGallerySource,IGalleryVM>(new GalleryVM(controlId)) as IControlVM;

        /// <summary>Returns a new Ribbon LabelControl view-model instance.</summary>
        internal LabelControlVM NewLabelControl(string controlId)
        => Add<LabelControlVM, ILabelControlSource,ILabelControlVM>(new LabelControlVM(controlId));

        /// <summary>Returns a new Ribbon BoxControl view-model instance.</summary>
        internal BoxControlVM NewBoxControl(string controlId, IEnumerable<IControlVM> controls)
        => Add<BoxControlVM, IBoxControlSource,IBoxControlVM>(new BoxControlVM(controlId, controls));

        /// <summary>Returns a new Ribbon LabelControl view-model instance.</summary>
        internal MenuSeparatorVM NewMenuSeparator(string controlId)
        => Add<MenuSeparatorVM, IMenuSeparatorSource,IMenuSeparatorVM>(new MenuSeparatorVM(controlId));

        /// <summary>Returns a new Ribbon Split(Toggle)Button view-model instance.</summary>
        internal SplitToggleButtonVM NewSplitToggleButton(string controlId, IMenuVM menu, IToggleVM toggle)
        => Add<SplitToggleButtonVM, IToggleSource,ISplitToggleButtonVM>(new SplitToggleButtonVM(this, controlId, menu, toggle));

        /// <summary>Returns a new Ribbon Split(Press)Button view-model instance.</summary>
        internal SplitPressButtonVM NewSplitPressButton(string controlId, IMenuVM menu, IButtonVM button)
        => Add<SplitPressButtonVM, IButtonSource,ISplitPressButtonVM>(new SplitPressButtonVM(this, controlId, menu, button));

        /// <summary>Returns a new Ribbon ToggleButton view-model instance.</summary>
        internal MenuVM NewMenu(string controlId, IEnumerable<IControlVM> controls)
        => Add<MenuVM, IMenuSource,IMenuVM>(new MenuVM(this, controlId, controls));

        /// <summary>Returns a new Ribbon ToggleButton view-model instance.</summary>
        internal DynamicMenuVM NewDynamicMenu(string controlId)
        => Add<DynamicMenuVM, IDynamicMenuSource,IDynamicMenuVM>(new DynamicMenuVM(this, controlId));
        #endregion
    }
}

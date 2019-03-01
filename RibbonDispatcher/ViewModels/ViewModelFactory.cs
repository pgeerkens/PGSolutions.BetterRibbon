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

namespace PGSolutions.RibbonDispatcher.ViewModels {
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
    [CLSCompliant(true)]
    [Description("Implementation of the factory for Ribbon objects. Visible to enable TypeLib creation.")]
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
        }

        /// <summary>.</summary>
        internal event ChangedEventHandler Changed;

        /// <summary>.</summary>
        internal void OnChanged(object sender, IControlChangedEventArgs e)
        => Changed?.Invoke(this, new ControlChangedEventArgs(e.ControlId));

        /// <summary>.</summary>
        internal TControl GetControl<TControl>(string controlId) where TControl : class, IControlVM
        => Controls.FirstOrDefault(c => c.Key == controlId).Value as TControl;

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
            if (!_controls.ContainsKey(ctrl.Id)) {
                _controls       .Add(ctrl.Id, ctrl);
                _editables      .AddNotNull(ctrl.Id, ctrl as IEditableVM);
                _sizeables      .AddNotNull(ctrl.Id, ctrl as ISizeableVM);
                _imageables     .AddNotNull(ctrl.Id, ctrl as IImageableVM);
                _clickables     .AddNotNull(ctrl.Id, ctrl as IClickableVM);
                _toggleables    .AddNotNull(ctrl.Id, ctrl as IToggleableVM);
                _selectItems    .AddNotNull(ctrl.Id, ctrl as ISelectItemsVM);
                _selectables    .AddNotNull(ctrl.Id, ctrl as ISelectablesVM);
                _dynamicMenus   .AddNotNull(ctrl.Id, ctrl as IDynamicMenuVM);
                _gallerySizes   .AddNotNull(ctrl.Id, ctrl as IGallerySizeVM);
                _menuSeparators .AddNotNull(ctrl.Id, ctrl as IMenuSeparatorVM);
                _descriptionable.AddNotNull(ctrl.Id, ctrl as IDescriptionableVM);

                ctrl.Changed += OnChanged;
            }
            return ctrl;
        }

        internal void Remove(IContainerControl requestor, IControlVM ctrl) {
            if (requestor==null  ||  ctrl == null) return;

            if (_descriptionable.ContainsKey(ctrl.Id)) _descriptionable.Remove(ctrl.Id);
            if (_menuSeparators .ContainsKey(ctrl.Id)) _menuSeparators .Remove(ctrl.Id);
            if (_gallerySizes   .ContainsKey(ctrl.Id)) _gallerySizes   .Remove(ctrl.Id);
            if (_dynamicMenus   .ContainsKey(ctrl.Id)) _dynamicMenus   .Remove(ctrl.Id);
            if (_selectables    .ContainsKey(ctrl.Id)) _selectables    .Remove(ctrl.Id);
            if (_selectItems    .ContainsKey(ctrl.Id)) _selectItems    .Remove(ctrl.Id);
            if (_toggleables    .ContainsKey(ctrl.Id)) _toggleables    .Remove(ctrl.Id);
            if (_clickables     .ContainsKey(ctrl.Id)) _clickables     .Remove(ctrl.Id);
            if (_imageables     .ContainsKey(ctrl.Id)) _imageables     .Remove(ctrl.Id);
            if (_sizeables      .ContainsKey(ctrl.Id)) _sizeables      .Remove(ctrl.Id);
            if (_editables      .ContainsKey(ctrl.Id)) _editables      .Remove(ctrl.Id);
            if (_controls       .ContainsKey(ctrl.Id)) _controls       .Remove(ctrl.Id);

            ctrl.OnPurged(requestor);
        }
        #endregion

        #region Factoy Method implementation
        /// <summary>.</summary>
        internal TabVM NewTab(string controlId)
        => Add<TabVM,IControlSource,ITabVM>(new TabVM(this, controlId));

        /// <summary>Returns a new Ribbon Group view-model instance.</summary>
        internal GroupVM NewGroup(string controlId)
        => Add<GroupVM,IControlSource,IGroupVM>(new GroupVM(this, controlId));

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
        internal DropDownVM NewDropDown(string controlId)
        => Add<DropDownVM,IDropDownSource,IDropDownVM>(new DropDownVM(controlId));

        /// <summary>Returns a new Ribbon DropDown view-model instance.</summary>
        internal StaticDropDownVM NewStaticDropDown(string controlId, IReadOnlyList<StaticItemVM> items)
        => Add<StaticDropDownVM,IStaticDropDownSource,IDropDownVM>(new StaticDropDownVM(controlId,items));

        /// <summary>Returns a new Ribbon SelectableItem view-model instance.</summary>
        [SuppressMessage("Microsoft.Performance","CA1822:MarkMembersAsStatic")]
        internal StaticItemVM NewStaticItem(string controlId, IControlStrings strings)
        => new StaticItemVM(controlId, strings);

        /// <summary>Returns a new Ribbon EditBox view-model instance.</summary>
        internal EditBoxVM NewEditBox(string controlId)
        => Add<EditBoxVM, IEditBoxSource,IEditBoxVM>(new EditBoxVM(controlId));

        /// <summary>Returns a new Ribbon ComboBox view-model instance.</summary>
        internal ComboBoxVM NewComboBox(string controlId)
        => Add<ComboBoxVM, IComboBoxSource,IComboBoxVM>(new ComboBoxVM(controlId));

        /// <summary>Returns a new Ribbon ComboBox view-model instance.</summary>
        internal StaticComboBoxVM NewStaticComboBox(string controlId, IReadOnlyList<StaticItemVM> items)
        => Add<StaticComboBoxVM, IStaticComboBoxSource,IStaticComboBoxVM>(new StaticComboBoxVM(controlId,items));

        /// <summary>Returns a new Ribbon ComboBox view-model instance.</summary>
        internal GalleryVM NewGallery(string controlId)
        => Add<GalleryVM, IGallerySource,IGalleryVM>(new GalleryVM(controlId));

        /// <summary>Returns a new Ribbon ComboBox view-model instance.</summary>
        internal StaticGalleryVM NewStaticGallery(string controlId, IReadOnlyList<StaticItemVM> items)
        => Add<StaticGalleryVM, IStaticGallerySource,IStaticGalleryVM>(new StaticGalleryVM(controlId,items));

        /// <summary>Returns a new Ribbon LabelControl view-model instance.</summary>
        internal LabelControlVM NewLabelControl(string controlId)
        => Add<LabelControlVM, ILabelControlSource,ILabelControlVM>(new LabelControlVM(controlId));

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
        internal MenuVM NewMenu(string controlId)
        => Add<MenuVM, IMenuSource,IMenuVM>(new MenuVM(this, controlId));
        #endregion
    }
}

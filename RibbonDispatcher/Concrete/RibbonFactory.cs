using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using stdole;

using Microsoft.Office.Core;

using PGSolutions.RibbonDispatcher.ControlMixins;
using PGSolutions.RibbonDispatcher.AbstractCOM;
using PGSolutions.RibbonDispatcher.Concrete;
using static PGSolutions.RibbonDispatcher.AbstractCOM.RdControlSize;

namespace PGSolutions.RibbonDispatcher {

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
    public class RibbonFactory : IRibbonFactory {
        internal RibbonFactory(IRibbonUI ribbonUI) : this(ribbonUI, new ResourceLoader(), null) { ; }

        internal RibbonFactory(IRibbonUI ribbonUI, IResourceManager manager) : this(ribbonUI, null, manager) { ; }

        internal RibbonFactory(IRibbonUI ribbonUI, ResourceLoader loader, IResourceManager manager) {
            _ribbonUI        = ribbonUI;
            ResourceLoader   = loader;
            ResourceManager  = manager ?? loader;

            _controls    = new Dictionary<string, IRibbonCommon>();
            _sizeables   = new Dictionary<string, ISizeableMixin>();
            _actionables = new Dictionary<string, IClickableMixin>();
            _toggleables = new Dictionary<string, IToggleableMixin>();
            _selectables = new Dictionary<string, ISelectableMixin>();
            _imageables  = new Dictionary<string, IImageableMixin>();
        }

        internal IResourceLoader  ResourceLoader  { get; }
        public   IResourceManager ResourceManager { get; }

        private  readonly IRibbonUI                             _ribbonUI;
        private  readonly IDictionary<string, IRibbonCommon>    _controls;
        private  readonly IDictionary<string, ISizeableMixin>   _sizeables;
        private  readonly IDictionary<string, IClickableMixin> _actionables;
        private  readonly IDictionary<string, ISelectableMixin> _selectables;
        private  readonly IDictionary<string, IImageableMixin>  _imageables;
        private  readonly IDictionary<string, IToggleableMixin> _toggleables;

        internal object LoadImage(string imageId) => ResourceManager.GetImage(imageId);

        /// <summary>Returns a readonly collection of all Ribbon Controls in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string, IRibbonCommon>    Controls    => new ReadOnlyDictionary<string, IRibbonCommon>(_controls);
 
        /// <summary>Returns a readonly collection of all Ribbon (Action) Buttons in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string, ISizeableMixin>   Sizeables   => new ReadOnlyDictionary<string, ISizeableMixin>(_sizeables);

        /// <summary>Returns a readonly collection of all Ribbon (Action) Buttons in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string, IClickableMixin> Actionables => new ReadOnlyDictionary<string, IClickableMixin>(_actionables);

        /// <summary>Returns a readonly collection of all Ribbon DropDowns in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string, ISelectableMixin> Selectables => new ReadOnlyDictionary<string, ISelectableMixin>(_selectables);

        /// <summary>Returns a readonly collection of all Ribbon Imageable Controls in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string, IImageableMixin>  Imageables  => new ReadOnlyDictionary<string, IImageableMixin>(_imageables);

        /// <summary>Returns a readonly collection of all Ribbon Toggle Buttons in this Ribbon ViewModel.</summary>
        internal IReadOnlyDictionary<string, IToggleableMixin> Toggleables => new ReadOnlyDictionary<string, IToggleableMixin>(_toggleables);

        private void PropertyChanged(object sender, IControlChangedEventArgs e) => _ribbonUI.InvalidateControl(e.ControlId);

        private T Add<T>(T ctrl) where T:RibbonCommon {
            _controls.Add(ctrl.Id, ctrl);

            _actionables.AddNotNull(ctrl.Id, ctrl as IClickableMixin);
            _sizeables.AddNotNull(ctrl.Id, ctrl as ISizeableMixin);
            _selectables.AddNotNull(ctrl.Id, ctrl as ISelectableMixin);
            _imageables.AddNotNull(ctrl.Id, ctrl as IImageableMixin);
            _toggleables.AddNotNull(ctrl.Id, ctrl as IToggleableMixin);

            ctrl.Changed += PropertyChanged;
            return ctrl;
        }

        /// <summary>TODO</summary>
        internal void Invalidate()                            => _ribbonUI.Invalidate();

        /// <summary>TODO</summary>
        internal void InvalidateControl(string ControlId)     => _ribbonUI.InvalidateControl(ControlId);

        /// <summary>TODO</summary>
        internal void InvalidateControlMso(string ControlId)  => _ribbonUI.InvalidateControlMso(ControlId);

        /// <summary>TODO</summary>
        internal void ActivateTab(string ControlId)           => _ribbonUI.ActivateTab(ControlId);

        /// <summary>Returns a new Ribbon Group ViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public RibbonGroup NewRibbonGroup(string ItemId, bool Visible = true, bool Enabled = true)
            => Add(new RibbonGroup(ItemId, ResourceManager, Visible, Enabled));

        /// <summary>Returns a new Ribbon ActionButton ViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public RibbonButton NewRibbonButton(string ItemId, bool Visible = true, bool Enabled = true,
            RdControlSize Size      = rdLarge,
            IPictureDisp  Image     = null,
            bool          ShowImage = true,
            bool          ShowLabel = true
        ) => Add(new RibbonButton(ItemId, ResourceManager, Visible, Enabled, Size, new ImageObject(Image), ShowImage, ShowLabel));

        /// <summary>Returns a new Ribbon ActionButton ViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public RibbonButton NewRibbonButtonMso(string ItemId, bool Visible = true, bool Enabled = true,
            RdControlSize Size      = rdLarge,
            string        ImageMso  = "Unknown",
            bool          ShowImage = true,
            bool          ShowLabel = true
        ) => Add(new RibbonButton(ItemId, ResourceManager, Visible, Enabled, Size, new ImageObject(ImageMso), ShowImage, ShowLabel));

        /// <summary>Returns a new Ribbon ToggleButton ViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public RibbonToggleButton NewRibbonToggle(string ItemId, bool Visible = true, bool Enabled = true,
            RdControlSize Size      = rdLarge,
            IPictureDisp  Image     = null,
            bool          ShowImage = true,
            bool          ShowLabel = true
        ) => Add(new RibbonToggleButton(ItemId, ResourceManager, Visible, Enabled, Size, new ImageObject(Image), ShowImage, ShowLabel));

        /// <summary>Returns a new Ribbon ToggleButton ViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public RibbonToggleButton NewRibbonToggleMso(string ItemId, bool Visible = true, bool Enabled = true,
            RdControlSize Size      = rdLarge,
            string        ImageMso  = "Unknown",
            bool          ShowImage = true,
            bool          ShowLabel = true
        ) => Add(new RibbonToggleButton(ItemId, ResourceManager, Visible, Enabled, Size, new ImageObject(ImageMso), ShowImage, ShowLabel));

        /// <summary>Returns a new Ribbon CheckBox ViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public RibbonCheckBox NewRibbonCheckBox(string ItemId, bool Visible = true, bool Enabled = true)
            => Add(new RibbonCheckBox(ItemId, ResourceManager, Visible, Enabled));

        /// <summary>Returns a new Ribbon DropDownViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public RibbonDropDown NewRibbonDropDown(string ItemId, bool Visible = true, bool Enabled = true)
            => Add(new RibbonDropDown(ItemId, ResourceManager, Visible, Enabled));

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public SelectableItem NewSelectableItem(string ItemId, IPictureDisp Image = null)
            => new SelectableItem(ItemId, ResourceManager, new ImageObject(Image));

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public SelectableItem NewSelectableItemMso(string ItemId, string ImageMso = "MacroSecurity")
            => new SelectableItem(ItemId, ResourceManager, new ImageObject(ImageMso));

        /// <inheritdoc/>
        public IResourceLoader NewResourceLoader() => ResourceLoader;
    }
}

using System.Runtime.CompilerServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ControlMixins {
    /// <summary>The mixin implementation for ISizeable ribbon controls.</summary>
    internal static class SizeableMixin {
        static ConditionalWeakTable<ISizeableMixin,Fields> _table = new ConditionalWeakTable<ISizeableMixin, Fields>();

        private sealed class Fields {
            public RdControlSize ControlSize = RdControlSize.rdLarge;
        }
        private static Fields Mixin(this ISizeableMixin sizeable) => _table.GetOrCreateValue(sizeable);

        /// <summary>Sets the {RdControlSize} value for an {ISizeableMixin} mixin.</summary>
        public static RdControlSize GetSize(this ISizeableMixin sizeable)
            => sizeable.Mixin().ControlSize;

        /// <summary>Sets the {RdControlSize} value for an {ISizeableMixin} mixin.</summary>
        public static void SetSize(this ISizeableMixin sizeable, RdControlSize size) {
            sizeable.Mixin().ControlSize = size;
            sizeable.OnChanged();
        }
    }
}

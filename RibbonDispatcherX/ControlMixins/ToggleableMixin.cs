////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.CompilerServices;

namespace PGSolutions.RibbonDispatcher.ControlMixins {
    /// <summary>Delegate type for </summary>
    [CLSCompliant(true)]
    public delegate void ToggledEventHandler(bool IsPressed);

    /// <summary>The mixin implementation for IToggleable ribbon controls.</summary>
    internal static class ToggleableMixin {
        static ConditionalWeakTable<IToggleableMixin, Fields> _table = new ConditionalWeakTable<IToggleableMixin, Fields>();

        private sealed class Fields {
            public bool IsPressed => Getter?.Invoke() ?? false;
            public Func<bool> Getter { private get; set; }
        }
        private static Fields Mixin(this IToggleableMixin mixin) => _table.GetOrCreateValue(mixin);

        public static void OnActionToggle(this IToggleableMixin mixin, bool isPressed) {
            mixin.OnToggled(isPressed);
            mixin.OnChanged();
        }

        public static void SetGetter (this IToggleableMixin mixin, Func<bool> getter) => mixin.Mixin().Getter = getter;
        public static bool GetPressed(this IToggleableMixin mixin) => mixin.Mixin().IsPressed;
        public static string GetLabel(this IToggleableMixin mixin)
            => mixin.Mixin().IsPressed && ! string.IsNullOrEmpty(mixin?.AlternateLabel())
                        ? mixin.AlternateLabel()
                        : mixin.LanguageStrings.Label ?? "";
        private static string AlternateLabel(this IToggleableMixin mixin) => mixin.LanguageStrings.AlternateLabel;
    }
}

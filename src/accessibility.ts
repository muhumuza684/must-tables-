"use strict";

export function makeFocusableButton(button: HTMLButtonElement, label: string): void {
    button.type = "button";
    button.setAttribute("aria-label", label);
    button.tabIndex = 0;
}

export function makeDialogAccessible(dialog: HTMLElement, label: string, labelledBy?: string): void {
    dialog.setAttribute("role", "dialog");
    dialog.setAttribute("aria-modal", "true");
    if (labelledBy) dialog.setAttribute("aria-labelledby", labelledBy);
    else dialog.setAttribute("aria-label", label);
}

export function trapFocus(container: HTMLElement, event: KeyboardEvent): void {
    if (event.key !== "Tab") return;
    const focusable = Array.from(container.querySelectorAll<HTMLElement>(
        'button:not([disabled]), [href], input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])'
    )).filter((element) => element.offsetParent !== null);
    if (focusable.length === 0) return;
    const first = focusable[0];
    const last = focusable[focusable.length - 1];
    if (event.shiftKey && document.activeElement === first) {
        event.preventDefault();
        last.focus();
    } else if (!event.shiftKey && document.activeElement === last) {
        event.preventDefault();
        first.focus();
    }
}

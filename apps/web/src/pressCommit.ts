/**
 * A press on a button while a number is being typed.
 *
 * A number is applied when its field loses focus (`change`, see controls.ts --
 * on purpose: running the model on every keystroke would run it on the
 * half-typed "19" of "1960"), and the press on a button is what takes the focus
 * from the field. So the page was redrawn between the button's mousedown and its
 * mouseup: the button the pointer went down on was gone, the one it came up on
 * was a new element, and a browser only clicks an element it saw go down. The
 * first press did nothing, the second -- with nothing left to apply -- worked. It
 * reached every button in a part of the page that is rebuilt on a change: the
 * tabs, Normalt/Avancerat, the language, the toggles and download buttons of the
 * results.
 *
 * Here a press on a button does not take the focus from the field (a mouse needs
 * no focus to click a button), and what was typed is applied when the click
 * arrives: the field is blurred, which fires its `change`, which redraws. If that
 * replaced the button, the click is handed on to the one that took its place --
 * the same button in the same place in the new page, and the one with the new
 * numbers behind it, which a download needs: it writes what was drawn, so it has
 * to be that button and not the one that was pressed.
 */

/** The kinds of `<input>` you type into; a checkbox or a range has nothing pending when it loses focus. */
const TYPED_INTO = new Set(["text", "number", "search", "tel", "url", "email"]);

/** What is needed of a button to tell it from the others; a real one has all of it. */
export type ButtonLike = Pick<HTMLButtonElement, "className" | "dataset" | "textContent" | "getAttribute">;

/**
 * What makes a button the same button once the page has been drawn again: what
 * kind it is, what its `data-*` say, what it is called. Not whether it is the
 * active one of a toggle -- that is the state the press is about to change.
 */
export function identityOf(button: ButtonLike): string {
  const kind = button.className.split(/\s+/).filter((name) => name !== "" && name !== "active");
  return JSON.stringify([
    kind.join(" "),
    { ...button.dataset },
    button.getAttribute("aria-label"),
    (button.textContent ?? "").trim(),
  ]);
}

/**
 * Which of the buttons now on the page is the one that was `before[pressed]`:
 * the one with its identity and the same number of identical ones in front of it
 * (the page has several "Ladda ner CSV", in an order that does not change with
 * the numbers). None if the redraw left no such button, in which case the press
 * is not handed on to a different one.
 */
export function counterpart(
  before: readonly string[],
  pressed: number,
  after: readonly string[],
): number | undefined {
  const identity = before[pressed];
  if (identity === undefined) return undefined;
  const wanted = before.slice(0, pressed).filter((other) => other === identity).length;
  let seen = 0;
  for (let i = 0; i < after.length; i += 1) {
    if (after[i] !== identity) continue;
    if (seen === wanted) return i;
    seen += 1;
  }
  return undefined;
}

function buttonOf(target: EventTarget | null): HTMLButtonElement | undefined {
  const button = target instanceof Element ? target.closest("button") : null;
  return button ?? undefined;
}

/** The field a number is being typed into, if the focus is in one. */
function fieldBeingTyped(): HTMLInputElement | undefined {
  const active = document.activeElement;
  return active instanceof HTMLInputElement && TYPED_INTO.has(active.type) ? active : undefined;
}

/** True while a click is being passed on, so that it is not taken for a new one. */
let handingOn = false;

/** Installs the two listeners; once, before the page is built. */
export function commitOnPress(root: Document = document): void {
  root.addEventListener(
    "mousedown",
    (event) => {
      if (event.button !== 0 || buttonOf(event.target) === undefined) return;
      if (fieldBeingTyped() === undefined) return;
      // The field keeps the focus, so nothing is applied -- and redrawn -- under the pointer.
      event.preventDefault();
    },
    true,
  );

  root.addEventListener(
    "click",
    (event) => {
      if (handingOn) return;
      const pressed = buttonOf(event.target);
      const field = fieldBeingTyped();
      if (pressed === undefined || field === undefined) return;

      const buttons = [...root.querySelectorAll("button")];
      const at = buttons.indexOf(pressed);
      const identities = buttons.map(identityOf);

      field.blur();
      if (pressed.isConnected) return;

      // The redraw replaced it. This click was for a button that is gone.
      event.stopImmediatePropagation();
      event.preventDefault();
      const now = [...root.querySelectorAll("button")];
      const next = now[counterpart(identities, at, now.map(identityOf)) ?? -1];
      if (next === undefined) return;
      handingOn = true;
      try {
        next.click();
      } finally {
        handingOn = false;
      }
    },
    true,
  );
}

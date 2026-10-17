/**
 * Which button a press is handed on to when the page is drawn again under it --
 * the part of `pressCommit.ts` that is arithmetic. The browser's side of it (a
 * press that lands, once, with the number that was just typed) is checked in
 * `tools/build/verify-offline.mjs` with a real mouse, since no test here has one.
 */
import { describe, expect, it } from "vitest";

import { counterpart, identityOf } from "../src/pressCommit.js";
import type { ButtonLike } from "../src/pressCommit.js";

const button = (patch: Partial<ButtonLike> = {}): ButtonLike => ({
  className: "export-btn",
  dataset: {},
  textContent: "Ladda ner CSV",
  getAttribute: () => null,
  ...patch,
});

describe("what makes a button the same button", () => {
  it("is its kind, its data-* and its text", () => {
    const csv = button({ dataset: { file: "tabell1.csv" } });
    expect(identityOf(csv)).toBe(identityOf(button({ dataset: { file: "tabell1.csv" } })));
    expect(identityOf(csv)).not.toBe(identityOf(button({ dataset: { file: "tabell2.csv" } })));
    expect(identityOf(csv)).not.toBe(identityOf(button({ dataset: { file: "tabell1.csv" }, textContent: "Ladda ner Excel" })));
    expect(identityOf(csv)).not.toBe(identityOf(button({ dataset: { file: "tabell1.csv" }, className: "panel-btn" })));
  });

  it("is not whether it is the active one of a toggle, which is what a press changes", () => {
    const off = button({ className: "panel-btn", dataset: { basis: "before-tax" } });
    const on = button({ className: "panel-btn active", dataset: { basis: "before-tax" } });
    expect(identityOf(on)).toBe(identityOf(off));
  });

  it("is not the spaces around its text or between its classes", () => {
    expect(identityOf(button({ textContent: "  Ladda ner CSV\n", className: "  export-btn  " }))).toBe(identityOf(button()));
  });

  it("tells a button with an aria-label from one without", () => {
    const labelled = button({ getAttribute: (name: string) => (name === "aria-label" ? "Stäng" : null) });
    expect(identityOf(labelled)).not.toBe(identityOf(button()));
  });
});

describe("the button a press is handed on to", () => {
  const csv = (file: string) => identityOf(button({ dataset: { file } }));
  const tab = (screen: string) => identityOf(button({ className: "panel-btn", dataset: { screen }, textContent: screen }));

  it("is the same button when the page has the same buttons", () => {
    const page = [tab("single"), tab("lasts"), csv("a.csv"), csv("b.csv")];
    for (let i = 0; i < page.length; i += 1) expect(counterpart(page, i, page)).toBe(i);
  });

  it("follows the button when others have come or gone around it", () => {
    const before = [tab("single"), csv("a.csv"), csv("b.csv")];
    const after = [tab("single"), tab("lasts"), csv("a.csv"), csv("b.csv"), csv("c.csv")];
    expect(counterpart(before, 2, after)).toBe(3);
    expect(counterpart(before, 0, after)).toBe(0);
  });

  it("takes the one in the same place among buttons that look alike", () => {
    const same = identityOf(button());
    const before = [same, tab("single"), same, same];
    const after = [tab("single"), same, same, same];
    expect(counterpart(before, 0, after)).toBe(1);
    expect(counterpart(before, 2, after)).toBe(2);
    expect(counterpart(before, 3, after)).toBe(3);
  });

  it("is none when the redraw left no such button -- never a different one", () => {
    const before = [tab("single"), csv("a.csv"), csv("b.csv")];
    expect(counterpart(before, 1, [tab("single"), csv("b.csv")])).toBeUndefined();
  });

  it("is none when fewer look alike than there were in front of the pressed one", () => {
    const same = identityOf(button());
    expect(counterpart([same, same, same], 2, [same, same])).toBeUndefined();
  });

  it("is none for a button that was not on the page", () => {
    const page = [tab("single")];
    expect(counterpart(page, -1, page)).toBeUndefined();
    expect(counterpart(page, 5, page)).toBeUndefined();
    expect(counterpart([], 0, page)).toBeUndefined();
  });
});

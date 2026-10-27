/**
 * "Om Pensionslyft": what the site is for, what it is not, and where to find the people behind it.
 *
 * Static text in two languages, no state and no model run. It is a screen of its own in the
 * main menu rather than a longer notice under the subtitle, so the header's short "Inofficiell
 * version" stays short and the whole disclaimer has somewhere to be read at length.
 */
import type { Lang } from "./i18n.js";
import { h, panel, say } from "./pensionAgeUi.js";

/**
 * Whose model this is and where it is published: Pensionsmyndigheten lets the tool be spread and modified on
 * condition that the agency is named as the source. `version` is the workbook's own ("ver. 4.8" on its first
 * sheet); it is the one thing here that moves with the yearly update (docs/YEARLY-UPDATE.md).
 */
export const SOURCE = {
  href: "https://www.pensionsmyndigheten.se/statistik-och-rapporter/pensionsmodellen/typfallsmodellen",
  text: "pensionsmyndigheten.se",
  version: "4.8",
} as const;

/** Where the site lives; the one piece of contact information there is. */
const WEBSITE = { href: "https://pensionslyft.se", text: "pensionslyft.se" } as const;

export interface AboutHandle {
  readonly element: HTMLElement;
  relabel(lang: Lang): void;
}

/** The disclaimer, one line each: what the page is not, and what it leaves to the reader. */
function disclaimer(l: Lang): readonly { readonly lead: string; readonly text: string }[] {
  return [
    {
      lead: say(l, "Inte finansiell rådgivning. ", "Not financial advice. "),
      text: say(
        l,
        "Innehållet och beräkningarna är allmän information. De är inte finansiell, juridisk, skattemässig eller annan rådgivning och ska inte ses som en rekommendation att göra, eller avstå från att göra, något.",
        "The content and the calculations are general information. They are not financial, legal, tax or other advice and are not a recommendation to do, or not to do, anything.",
      ),
    },
    {
      lead: say(l, "Prognoser, inte besked. ", "Forecasts, not statements. "),
      text: say(
        l,
        "Resultaten är beräkningar under de antaganden du själv anger, till exempel inflation, tillväxt, avkastning och livslängd. De är ingen förutsägelse och ingen garanti. Verkligheten kan avvika och regler kan ändras.",
        "The results are calculations under the assumptions you enter yourself, such as inflation, growth, return and lifespan. They are no prediction and no guarantee. Reality can differ and rules can change.",
      ),
    },
    {
      lead: say(l, "Förenklingar. ", "Simplifications. "),
      text: say(
        l,
        "Modellen är en förenkling av verkligheten. Det gäller även det privata sparandet, som räknas med förenklade skatteregler.",
        "The model is a simplification of reality. That includes private savings, which are calculated with simplified tax rules.",
      ),
    },
    {
      lead: say(l, "Inofficiell. ", "Unofficial. "),
      text: say(
        l,
        "Sidan är inte utvecklad av, kopplad till eller godkänd av Pensionsmyndigheten. Modellen, dess data och dess användarmanual är deras. Eventuella fel i den här versionen är inte Pensionsmyndighetens.",
        "This page is not built by, affiliated with or endorsed by Pensionsmyndigheten, the Swedish Pensions Agency. The model, its data and its user manual are theirs. Any errors in this version are not the agency's.",
      ),
    },
    {
      lead: say(l, "Eget ansvar. ", "At your own risk. "),
      text: say(
        l,
        "Du använder sidan på eget ansvar. Pensionslyft ansvarar, så långt lagen tillåter, inte för beslut som fattas eller förluster som uppstår med stöd av innehållet.",
        "You use the page at your own risk. To the extent the law allows, Pensionslyft is not responsible for decisions made or losses incurred on the basis of the content.",
      ),
    },
    {
      lead: say(l, "Kontrollera själv. ", "Check for yourself. "),
      text: say(
        l,
        "Jämför med din egen pensionsprognos på minPension.se och med din tjänstepensionsleverantör, och vänd dig vid behov till en oberoende rådgivare innan du fattar ett beslut.",
        "Compare with your own pension forecast at minPension.se and with your occupational pension provider, and turn to an independent adviser if needed before you decide.",
      ),
    },
    {
      lead: say(l, "Integritet. ", "Privacy. "),
      text: say(
        l,
        "Alla beräkningar sker i din webbläsare. Det du skriver in skickas inte någonstans.",
        "All calculations happen in your browser. What you enter is not sent anywhere.",
      ),
    },
  ];
}

export function createAboutPanel(lang: Lang): AboutHandle {
  const element = h("div", "about-panel");
  element.dataset.role = "about-panel";

  function build(l: Lang): void {
    const { section, body } = panel(say(l, "Om Pensionslyft", "About Pensionslyft"));

    body.append(
      h("h3", undefined, say(l, "Syfte", "Purpose")),
      h(
        "p",
        undefined,
        say(
          l,
          "Pensionslyft finns för att förmedla kunskap, inspirera och lyfta pensionsfrågan, så att du har bättre möjlighet att fatta välinformerade beslut som passar dig och din situation.",
          "Pensionslyft exists to share knowledge, inspire and bring the pension question to the fore, so that you are better able to make well-informed decisions that suit you and your situation.",
        ),
      ),
      h("h3", undefined, say(l, "Källa och bearbetning", "Source and adaptation")),
    );
    const source = h("p");
    const sourceLink = h("a", undefined, SOURCE.text);
    sourceLink.href = SOURCE.href;
    sourceLink.target = "_blank";
    sourceLink.rel = "noopener noreferrer";
    source.append(
      say(
        l,
        `Beräkningarna bygger på Typfallsmodellen (ver. ${SOURCE.version}), som Pensionsmyndigheten har tagit fram och publicerat. Källa: Pensionsmyndigheten, `,
        `The calculations are based on the Typfallsmodellen (ver. ${SOURCE.version}), which Pensionsmyndigheten, the Swedish Pensions Agency, has produced and published. Source: Pensionsmyndigheten, `,
      ),
      sourceLink,
      say(
        l,
        ". Pensionsmyndigheten tillåter att verktyget sprids och bearbetas om myndigheten anges som källa. Det här är en bearbetad version: modellen är omskriven för webben och har fått egna tillägg. Flikarna Jämför scenarier, Mikrosim, När ta ut pensionen?, Hur länge räcker pensionen? och Hur mycket får jag?, liksom det privata sparandet, är inte en del av Pensionsmyndighetens modell.",
        ". Pensionsmyndigheten allows the tool to be spread and modified as long as the agency is named as the source. This is an adapted version: the model has been rewritten for the web and given additions of its own. The tabs Compare scenarios, Mikrosim, When to take out the pension?, How long does the pension last? and How much do I get?, and the private savings, are not part of Pensionsmyndigheten's model.",
      ),
    );
    body.append(source, h("h3", undefined, say(l, "Friskrivning", "Disclaimer")));
    const list = h("ul");
    for (const item of disclaimer(l)) {
      const entry = h("li");
      entry.append(h("strong", undefined, item.lead), item.text);
      list.append(entry);
    }
    body.append(list, h("h3", undefined, say(l, "Kontakt", "Contact")));

    const contact = h("p");
    const link = h("a", undefined, WEBSITE.text);
    link.href = WEBSITE.href;
    link.target = "_blank";
    link.rel = "noopener noreferrer";
    contact.append(say(l, "Webbplats: ", "Website: "), link);
    body.append(contact);

    element.replaceChildren(section);
  }

  build(lang);
  return { element, relabel: build };
}

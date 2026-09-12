/**
 * Jobbskatteavdrag -- the earned income tax credit, one function per rule year.
 *
 * Translated mechanically alongside the grundavdrag functions; see the note in
 * grundavdrag.ts.
 *
 * Each computes a credit from earned income, nets off the grundavdrag that
 * would otherwise apply to it, and multiplies by the municipal tax rate --
 * which is why every one of them calls `avdragxx`.
 *
 * QUIRK worth knowing: `Jobb07` and `Jobb08` declare their working variable
 * `As Long`, so every assignment inside them rounds to a whole krona. The other
 * ten use `As Double` and do not. The translator preserves this by wrapping
 * those assignments in `vbaCLng`.
 */

import { vbaCLng, vbaInt } from "../vba/math.js";
import { avdragxx } from "./grundavdrag.js";

export function Jobb07(inkomst: number, alder = 30, ksats = 0.3144, pbb = 42400, marginal = 0, binkomst = 0, year = 2100, wyear = 21000, IBB = 0, kvoten = 1, Iyear = 0): number {
  let result = 0;
  let avdrag = 0;
  avdrag = avdragxx(inkomst + binkomst, pbb, marginal, alder, year, wyear, IBB, kvoten, Iyear);
  let jobb = 0;
  if (inkomst <= 0.79 * pbb) {
    jobb = vbaCLng(inkomst);
  } else if (inkomst <= 2.72 * pbb) {
    jobb = vbaCLng(0.2 * (inkomst - 0.79 * pbb) + 0.79 * pbb);
  } else {
    jobb = vbaCLng(1.176 * pbb);
  }
  if (alder >= 66) {
    if (inkomst <= 1.59 * pbb) {
      jobb = vbaCLng(inkomst);
    } else if (inkomst <= 2.72 * pbb) {
      jobb = vbaCLng(1.5 * pbb + 0.2 * (inkomst - 1.59 * pbb));
    } else {
      jobb = vbaCLng(1.816 * pbb);
    }
  }
  jobb = vbaCLng((jobb - avdrag) * ksats);
  if (jobb < 0) jobb = vbaCLng(0);
  if (marginal === 0) jobb = vbaCLng(vbaInt(jobb));
  result = jobb;
  return result;
}

export function Jobb08(inkomst: number, alder = 30, ksats = 0.3144, pbb = 42400, marginal = 0, binkomst = 0, year = 2100, wyear = 21000, IBB = 0, kvoten = 1, Iyear = 0): number {
  let result = 0;
  let avdrag = 0;
  avdrag = avdragxx(inkomst + binkomst, pbb, marginal, alder, year, wyear, IBB, kvoten, Iyear);
  let jobb = 0;
  if (inkomst <= 0.91 * pbb) {
    jobb = vbaCLng(inkomst);
  } else if (inkomst <= 2.72 * pbb) {
    jobb = vbaCLng(0.2 * (inkomst - 0.91 * pbb) + 0.91 * pbb);
  } else if (inkomst <= 7 * pbb) {
    jobb = vbaCLng(0.033 * (inkomst - 2.72 * pbb) + 1.272 * pbb);
  } else {
    jobb = vbaCLng(1.413 * pbb);
  }
  if (alder >= 66) {
    if (inkomst <= 1.79 * pbb) {
      jobb = vbaCLng(inkomst);
    } else if (inkomst <= 2.72 * pbb) {
      jobb = vbaCLng(1.79 * pbb + 0.2 * (inkomst - 1.79 * pbb));
    } else {
      jobb = vbaCLng(2.117 * pbb);
    }
  }
  jobb = vbaCLng((jobb - avdrag) * ksats);
  if (jobb < 0) jobb = vbaCLng(0);
  if (marginal === 0) jobb = vbaCLng(vbaInt(jobb));
  result = jobb;
  return result;
}

export function Jobb10(inkomst: number, alder = 30, ksats = 0.3144, pbb = 42400, marginal = 0, binkomst = 0, year = 2100, wyear = 21000, IBB = 0, kvoten = 1, Iyear = 0): number {
  let result = 0;
  let avdrag = 0;
  avdrag = avdragxx(inkomst + binkomst, pbb, marginal, alder, year, wyear, IBB, kvoten, Iyear);
  let jobb = 0;
  if (inkomst <= 0.91 * pbb) {
    jobb = vbaCLng(inkomst);
  } else if (inkomst <= 2.72 * pbb) {
    jobb = vbaCLng(0.304 * (inkomst - 0.91 * pbb) + 0.91 * pbb);
  } else if (inkomst <= 7 * pbb) {
    jobb = vbaCLng(1.461 * pbb + 0.095 * (inkomst - 2.72 * pbb));
  } else {
    jobb = vbaCLng(1.868 * pbb);
  }
  jobb = vbaCLng((jobb - avdrag) * ksats);
  if (alder >= 66) {
    if (inkomst <= 100000) {
      jobb = vbaCLng(0.2 * inkomst);
    } else if (inkomst <= 300000) {
      jobb = vbaCLng(15000 + 0.05 * inkomst);
    } else {
      jobb = vbaCLng(30000);
    }
  }
  if (jobb < 0) jobb = vbaCLng(0);
  if (marginal === 0) jobb = vbaCLng(vbaInt(jobb));
  result = jobb;
  return result;
}

export function Jobb11(inkomst: number, alder = 30, ksats = 0.3144, pbb = 42400, marginal = 0, binkomst = 0, year = 2100, wyear = 21000, IBB = 0, kvoten = 1, Iyear = 0): number {
  let result = 0;
  let avdrag = 0;
  let jobb = 0;
  avdrag = avdragxx(inkomst + binkomst, pbb, marginal, alder, year, wyear, IBB, kvoten, Iyear);
  if (inkomst <= 0.91 * pbb) {
    jobb = inkomst;
  } else if (inkomst <= 2.72 * pbb) {
    jobb = 0.304 * (inkomst - 0.91 * pbb) + 0.91 * pbb;
  } else if (inkomst <= 7 * pbb) {
    jobb = 1.461 * pbb + 0.095 * (inkomst - 2.72 * pbb);
  } else {
    jobb = 1.868 * pbb;
  }
  jobb = (jobb - avdrag) * ksats;
  if (alder >= 66) {
    if (inkomst <= 100000) {
      jobb = 0.2 * inkomst;
    } else if (inkomst <= 300000) {
      jobb = 15000 + 0.05 * inkomst;
    } else {
      jobb = 30000;
    }
  }
  result = jobb;
  if (result < 0) result = 0;
  if (marginal === 0) result = vbaInt(result);
  return result;
}

export function Jobb14(inkomst: number, alder = 30, ksats = 0.3144, pbb = 42400, marginal = 0, binkomst = 0, year = 2100, wyear = 21000, IBB = 0, kvoten = 1, Iyear = 0): number {
  let result = 0;
  let avdrag = 0;
  let jobb = 0;
  avdrag = avdragxx(inkomst + binkomst, pbb, marginal, alder, year, wyear, IBB, kvoten, Iyear);
  if (inkomst <= 0.91 * pbb) {
    jobb = inkomst;
  } else if (inkomst <= 2.94 * pbb) {
    jobb = 0.332 * (inkomst - 0.91 * pbb) + 0.91 * pbb;
  } else if (inkomst <= 8.08 * pbb) {
    jobb = 1.584 * pbb + 0.111 * (inkomst - 2.94 * pbb);
  } else {
    jobb = 2.155 * pbb;
  }
  jobb = (jobb - avdrag) * ksats;
  if (alder >= 66) {
    if (inkomst <= 100000) {
      jobb = 0.2 * inkomst;
    } else if (inkomst <= 300000) {
      jobb = 15000 + 0.05 * inkomst;
    } else {
      jobb = 30000;
    }
  }
  result = jobb;
  if (result < 0) result = 0;
  if (marginal === 0) result = vbaInt(result);
  return result;
}

export function Jobb16(inkomst: number, alder = 30, ksats = 0.3144, pbb = 42800, marginal = 0, binkomst = 0, year = 2100, wyear = 21000, IBB = 0, kvoten = 1, Iyear = 0): number {
  let result = 0;
  let avdrag = 0;
  let jobb = 0;
  avdrag = avdragxx(inkomst + binkomst, pbb, marginal, alder, year, wyear, IBB, kvoten, Iyear);
  if (inkomst <= 0.91 * pbb) {
    jobb = inkomst;
  } else if (inkomst <= 2.94 * pbb) {
    jobb = 0.332 * (inkomst - 0.91 * pbb) + 0.91 * pbb;
  } else if (inkomst <= 8.08 * pbb) {
    jobb = 1.584 * pbb + 0.111 * (inkomst - 2.94 * pbb);
  } else {
    jobb = 2.155 * pbb;
  }
  if (inkomst <= 13.54 * pbb) {
    jobb = (jobb - avdrag) * ksats;
  } else {
    jobb = (jobb - avdrag) * ksats - 0.03 * (inkomst - 13.54 * pbb);
  }
  if (alder >= 66) {
    if (inkomst <= 100000) {
      jobb = 0.2 * inkomst;
    } else if (inkomst <= 300000) {
      jobb = 15000 + 0.05 * inkomst;
    } else if (inkomst <= 600000) {
      jobb = 30000;
    } else {
      jobb = 30000 - 0.05 * (inkomst - 600000);
    }
  }
  result = jobb;
  if (result < 0) result = 0;
  if (marginal === 0) result = vbaInt(result);
  return result;
}

export function Jobb19(inkomst: number, alder = 30, ksats = 0.3144, pbb = 46500, marginal = 0, binkomst = 0, year = 2100, wyear = 21000, IBB = 0, kvoten = 1, Iyear = 0, Xage = 66): number {
  let result = 0;
  let avdrag = 0;
  let jobb = 0;
  avdrag = avdragxx(inkomst + binkomst, pbb, marginal, alder, year, wyear, IBB, kvoten, Iyear);
  if (inkomst <= 0.91 * pbb) {
    jobb = inkomst;
  } else if (inkomst <= 3.24 * pbb) {
    jobb = 0.3405 * (inkomst - 0.91 * pbb) + 0.91 * pbb;
  } else if (inkomst <= 8.08 * pbb) {
    jobb = 1.703 * pbb + 0.128 * (inkomst - 3.24 * pbb);
  } else {
    jobb = 2.323 * pbb;
  }
  if (inkomst <= 13.54 * pbb) {
    jobb = (jobb - avdrag) * ksats;
  } else {
    jobb = (jobb - avdrag) * ksats - 0.03 * (inkomst - 13.54 * pbb);
  }
  if (alder >= Xage) {
    if (inkomst <= 100000) {
      jobb = 0.2 * inkomst;
    } else if (inkomst <= 300000) {
      jobb = 15000 + 0.05 * inkomst;
    } else if (inkomst <= 600000) {
      jobb = 30000;
    } else {
      jobb = 30000 - 0.05 * (inkomst - 600000);
    }
  }
  if (jobb < 0) jobb = 0;
  if (marginal === 0) jobb = vbaInt(jobb);
  result = jobb;
  return result;
}

export function Jobb22(inkomst: number, alder = 30, ksats = 0.3144, pbb = 46500, marginal = 0, binkomst = 0, year = 2100, wyear = 21000, IBB = 0, kvoten = 1, Iyear = 0, Xage = 66): number {
  let result = 0;
  let avdrag = 0;
  let jobb = 0;
  avdrag = avdragxx(inkomst + binkomst, pbb, marginal, alder, year, wyear, IBB, kvoten, Iyear);
  if (inkomst <= 0.91 * pbb) {
    jobb = inkomst;
  } else if (inkomst <= 3.24 * pbb) {
    jobb = 0.3874 * (inkomst - 0.91 * pbb) + 0.91 * pbb;
  } else if (inkomst <= 8.08 * pbb) {
    jobb = 1.812 * pbb + 0.128 * (inkomst - 3.24 * pbb);
  } else {
    jobb = 2.432 * pbb;
  }
  if (inkomst <= 13.54 * pbb) {
    jobb = (jobb - avdrag) * ksats;
  } else {
    jobb = (jobb - avdrag) * ksats - 0.03 * (inkomst - 13.54 * pbb);
  }
  if (alder >= Xage) {
    if (inkomst <= 100000) {
      jobb = 0.2 * inkomst;
    } else if (inkomst <= 300000) {
      jobb = 15000 + 0.05 * inkomst;
    } else if (inkomst <= 600000) {
      jobb = 30000;
    } else {
      jobb = 30000 - 0.05 * (inkomst - 600000);
    }
  }
  if (jobb < 0) jobb = 0;
  if (marginal === 0) jobb = vbaInt(jobb);
  result = jobb;
  return result;
}

export function Jobb23(inkomst: number, alder = 30, ksats = 0.3144, pbb = 46500, marginal = 0, binkomst = 0, year = 2100, wyear = 21000, IBB = 0, kvoten = 1, Iyear = 0, Xage = 66): number {
  let result = 0;
  let avdrag = 0;
  let jobb = 0;
  avdrag = avdragxx(inkomst + binkomst, pbb, marginal, alder, year, wyear, IBB, kvoten, Iyear);
  if (inkomst <= 0.91 * pbb) {
    jobb = inkomst;
  } else if (inkomst <= 3.24 * pbb) {
    jobb = 0.3874 * (inkomst - 0.91 * pbb) + 0.91 * pbb;
  } else if (inkomst <= 8.08 * pbb) {
    jobb = 1.812 * pbb + 0.128 * (inkomst - 3.24 * pbb);
  } else {
    jobb = 2.432 * pbb;
  }
  if (inkomst <= 13.54 * pbb) {
    jobb = (jobb - avdrag) * ksats;
  } else {
    jobb = (jobb - avdrag) * ksats - 0.03 * (inkomst - 13.54 * pbb);
  }
  if (alder >= 66) {
    if (inkomst <= 100000) {
      jobb = 0.22 * inkomst;
    } else if (inkomst <= 300000) {
      jobb = 15000 + 0.07 * inkomst;
    } else if (inkomst <= 600000) {
      jobb = 36000;
    } else {
      jobb = 36000 - 0.03 * (inkomst - 600000);
    }
  }
  if (jobb < 0) jobb = 0;
  if (marginal === 0) jobb = vbaInt(jobb);
  result = jobb;
  return result;
}

export function Jobb24(inkomst: number, alder = 30, ksats = 0.3144, pbb = 46500, marginal = 0, binkomst = 0, year = 2100, wyear = 21000, IBB = 0, kvoten = 1, Iyear = 0, Xage = 66): number {
  let result = 0;
  let avdrag = 0;
  let jobb = 0;
  avdrag = avdragxx(inkomst + binkomst, pbb, marginal, alder, year, wyear, IBB, kvoten, Iyear);
  if (inkomst <= 0.91 * pbb) {
    jobb = inkomst;
  } else if (inkomst <= 3.24 * pbb) {
    jobb = 0.3874 * (inkomst - 0.91 * pbb) + 0.91 * pbb;
  } else if (inkomst <= 8.08 * pbb) {
    jobb = 1.813 * pbb + 0.1643 * (inkomst - 3.24 * pbb);
  } else if (inkomst <= 13.54 * pbb) {
    jobb = 2.608 * pbb;
  } else {
    jobb = 2.608 * pbb - 0.03 * (inkomst - 13.54 * pbb);
  }
  jobb = (jobb - avdrag) * ksats;
  if (alder > Xage) {
    if (inkomst < 1.75 * pbb) {
      jobb = 0.22 * inkomst;
    } else if (inkomst < 5.24 * pbb) {
      jobb = 0.2635 * pbb + 0.07 * inkomst;
    } else if (inkomst < 10.48 * pbb) {
      jobb = 0.6293 * pbb;
    } else {
      jobb = 0.6293 * pbb - 0.03 * (inkomst - 10.48 * pbb);
    }
  }
  if (jobb < 0) jobb = 0;
  if (marginal === 0) jobb = vbaInt(jobb);
  result = jobb;
  return result;
}

export function Jobb25(inkomst: number, alder = 30, ksats = 0.3144, pbb = 46500, marginal = 0, binkomst = 0, year = 2100, wyear = 21000, IBB = 0, kvoten = 1, Iyear = 0, Xage = 66): number {
  let result = 0;
  let avdrag = 0;
  let jobb = 0;
  avdrag = avdragxx(inkomst + binkomst, pbb, marginal, alder, year, wyear, IBB, kvoten, Iyear);
  if (inkomst <= 0.91 * pbb) {
    jobb = inkomst;
  } else if (inkomst <= 3.24 * pbb) {
    jobb = 0.3874 * (inkomst - 0.91 * pbb) + 0.91 * pbb;
  } else if (inkomst <= 8.08 * pbb) {
    jobb = 1.813 * pbb + 0.199 * (inkomst - 3.24 * pbb);
  } else {
    jobb = 2.776 * pbb;
  }
  jobb = (jobb - avdrag) * ksats;
  if (alder > Xage) {
    if (inkomst < 1.7 * pbb) {
      jobb = 0.22 * inkomst;
    } else if (inkomst < 5.24 * pbb) {
      jobb = 0.2635 * pbb + 0.07 * inkomst;
    } else {
      jobb = 0.6293 * pbb;
    }
  }
  if (jobb < 0) jobb = 0;
  if (marginal === 0) jobb = vbaInt(jobb);
  result = jobb;
  return result;
}

export function Jobb26(inkomst: number, alder = 30, ksats = 0.3144, pbb = 46500, marginal = 0, binkomst = 0, year = 2100, wyear = 21000, IBB = 0, kvoten = 1, Iyear = 0, Xage = 66): number {
  let result = 0;
  let avdrag = 0;
  let jobb = 0;
  avdrag = avdragxx(inkomst + binkomst, pbb, marginal, alder, year, wyear, IBB, kvoten, Iyear, Xage);
  if (inkomst <= 0.91 * pbb) {
    jobb = inkomst;
  } else if (inkomst <= 3.24 * pbb) {
    jobb = 0.3874 * (inkomst - 0.91 * pbb) + 0.91 * pbb;
  } else if (inkomst <= 8.08 * pbb) {
    jobb = 1.813 * pbb + 0.251 * (inkomst - 3.24 * pbb);
  } else {
    jobb = 3.027 * pbb;
  }
  jobb = (jobb - avdrag) * ksats;
  if (alder > Xage) {
    if (inkomst < 1.7 * pbb) {
      jobb = 0.22 * inkomst;
    } else if (inkomst < 6.5 * pbb) {
      jobb = 0.2635 * pbb + 0.07 * inkomst;
    } else {
      jobb = 0.6293 * pbb;
    }
  }
  if (jobb < 0) jobb = 0;
  if (marginal === 0) jobb = vbaInt(jobb);
  result = jobb;
  return result;
}


/**
 * Picks the jobbskatteavdrag rules for a year. Mirrors `Jobbxx`.
 *
 * There was no jobbskatteavdrag before 2007.
 */
export function Jobbxx(
  inkomst: number,
  alder = 30,
  ksats = 0.3144,
  pbb = 42400,
  marginal = 0,
  binkomst = 0,
  year = 2100,
  wyear = 21000,
  IBB = 0,
  kvoten = 1,
  IyearIn = 0,
  xage = 66,
): number {
  const Iyear = IyearIn === 0 ? year : IyearIn;
  const a = [inkomst, alder, ksats, pbb, marginal, binkomst, year, wyear, IBB, kvoten, Iyear] as const;

  if (Iyear < 2007) return 0;
  if (Iyear === 2007) return Jobb07(...a);
  if (Iyear === 2008) return Jobb08(...a);
  if (Iyear <= 2010) return Jobb10(...a);
  if (Iyear === 2011) return Jobb11(...a);
  if (Iyear < 2015) return Jobb14(...a);
  if (Iyear < 2019) return Jobb16(...a);
  if (Iyear < 2022) return Jobb19(...a);
  if (Iyear === 2022) return Jobb22(...a, xage);
  if (Iyear === 2023) return Jobb23(...a, xage);
  if (Iyear === 2024) return Jobb24(...a, xage);
  if (Iyear === 2025) return Jobb25(...a, xage);
  return Jobb26(...a, xage);
}

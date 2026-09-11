/**
 * Arrays indexed by age rather than from zero.
 *
 * The VBA works almost entirely in `ReDim x(startage To slutage)` arrays --
 * age 15 to 105 -- and reads `x(age - 1)` freely. Mirroring that indexing keeps
 * the ported code line-comparable with the original; rebasing to zero would
 * mean adjusting an offset on every one of several hundred accesses, which is
 * exactly the kind of edit that introduces an off-by-one nobody can find.
 *
 * Reads and writes outside the declared bounds throw, rather than silently
 * yielding `undefined` as a plain array would.
 */
export class AgeArray {
  readonly firstAge: number;
  readonly lastAge: number;
  private readonly values: Float64Array;

  constructor(firstAge: number, lastAge: number, fill = 0) {
    if (lastAge < firstAge) {
      throw new RangeError(`AgeArray: lastAge ${lastAge} is below firstAge ${firstAge}`);
    }
    this.firstAge = firstAge;
    this.lastAge = lastAge;
    this.values = new Float64Array(lastAge - firstAge + 1);
    if (fill !== 0) this.values.fill(fill);
  }

  get(age: number): number {
    const index = age - this.firstAge;
    if (index < 0 || index >= this.values.length) {
      throw new RangeError(`AgeArray: age ${age} outside ${this.firstAge}..${this.lastAge}`);
    }
    return this.values[index]!;
  }

  set(age: number, value: number): void {
    const index = age - this.firstAge;
    if (index < 0 || index >= this.values.length) {
      throw new RangeError(`AgeArray: age ${age} outside ${this.firstAge}..${this.lastAge}`);
    }
    this.values[index] = value;
  }

  /** Value at `age`, or 0 when outside the bounds -- for `x(age - 1)` at the lower edge. */
  getOrZero(age: number): number {
    const index = age - this.firstAge;
    return index < 0 || index >= this.values.length ? 0 : this.values[index]!;
  }

  add(age: number, delta: number): void {
    this.set(age, this.get(age) + delta);
  }

  /** Values for `firstAge..lastAge` inclusive, as a plain array. */
  slice(firstAge = this.firstAge, lastAge = this.lastAge): number[] {
    const out: number[] = [];
    for (let age = firstAge; age <= lastAge; age += 1) out.push(this.get(age));
    return out;
  }

  *entries(): Generator<[age: number, value: number]> {
    for (let age = this.firstAge; age <= this.lastAge; age += 1) {
      yield [age, this.get(age)];
    }
  }
}

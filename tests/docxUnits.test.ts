import { describe, it, expect } from "vitest";
import {
  twipsToPoints,
  pointsToTwips,
  twipsToInches,
  twipsToPx,
  pxToTwips,
  halfPointsToPoints,
  pointsToHalfPoints,
  eighthPointsToPoints,
  emuToInches,
  emuToPx,
  pxToEmu,
  emuToPoints,
  fiftiethsToPercent,
  percentToFiftieths,
  lineAutoToMultiple,
  parseMeasure,
} from "../src/units";

describe("units: OOXML measurement conversions", () => {
  it("twips ↔ points/inches/px around the 1-inch anchor", () => {
    expect(twipsToInches(1440)).toBe(1);
    expect(twipsToPoints(1440)).toBe(72);
    expect(twipsToPx(1440)).toBe(96); // 96 dpi
    expect(twipsToPoints(240)).toBe(12);
    expect(pointsToTwips(12)).toBe(240);
    expect(pxToTwips(96)).toBe(1440);
  });

  it("half-points are font size (w:sz val=24 → 12pt)", () => {
    expect(halfPointsToPoints(24)).toBe(12);
    expect(pointsToHalfPoints(11)).toBe(22);
  });

  it("eighth-points are border widths (val=8 → 1pt)", () => {
    expect(eighthPointsToPoints(8)).toBe(1);
    expect(eighthPointsToPoints(4)).toBe(0.5);
  });

  it("EMU around the 1-inch anchor", () => {
    expect(emuToInches(914400)).toBe(1);
    expect(emuToPoints(914400)).toBe(72);
    expect(emuToPx(914400)).toBe(96);
    expect(pxToEmu(96)).toBe(914400);
  });

  it("table-width percent is 1/50 of a percent (5000 → 100%)", () => {
    expect(fiftiethsToPercent(5000)).toBe(100);
    expect(fiftiethsToPercent(2500)).toBe(50);
    expect(percentToFiftieths(100)).toBe(5000);
  });

  it("auto line spacing is a 240ths multiple (360 → 1.5×)", () => {
    expect(lineAutoToMultiple(240)).toBe(1);
    expect(lineAutoToMultiple(360)).toBe(1.5);
    expect(lineAutoToMultiple(480)).toBe(2);
  });

  it("parseMeasure tolerates absent/garbage values", () => {
    expect(parseMeasure("240")).toBe(240);
    expect(parseMeasure("  240 ")).toBe(240);
    expect(parseMeasure(undefined)).toBe(0);
    expect(parseMeasure(null)).toBe(0);
    expect(parseMeasure("", 7)).toBe(7);
    expect(parseMeasure("abc", 7)).toBe(7);
    expect(parseMeasure("-15")).toBe(-15);
  });
});


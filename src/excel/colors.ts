import type ExcelJS from 'exceljs';
import type { ResolvedStyle } from '../domain/types';

interface ExcelColor {
  argb?: string;
  theme?: number;
  tint?: number;
}

function normalizeArgb(argb: string): string {
  const value = argb.replace(/^#/, '').toUpperCase();
  return value.length === 8 ? `#${value.slice(2)}` : `#${value}`;
}

function applyTint(channel: number, tint: number): number {
  if (tint < 0) return Math.round(channel * (1 + tint));
  return Math.round(channel * (1 - tint) + 255 * tint);
}

export function tintColor(hex: string, tint = 0): string {
  const value = hex.replace('#', '');
  const channels = [0, 2, 4].map((index) => Number.parseInt(value.slice(index, index + 2), 16));
  return `#${channels
    .map((channel) =>
      applyTint(channel ?? 0, tint)
        .toString(16)
        .padStart(2, '0'),
    )
    .join('')
    .toUpperCase()}`;
}

export function resolveExcelColor(
  color: ExcelColor | undefined,
  themeColors: string[],
): string | undefined {
  if (!color) return undefined;
  if (color.argb) return normalizeArgb(color.argb);
  if (typeof color.theme === 'number') {
    const base = themeColors[color.theme];
    return base ? tintColor(base, color.tint ?? 0) : undefined;
  }
  return undefined;
}

interface StyleCell {
  styleId?: number;
  fill?: ExcelJS.Fill;
  font?: Partial<ExcelJS.Font>;
}

export function resolveCellStyle(
  cell: ExcelJS.Cell,
  themeColors: string[],
  styleId?: number,
): ResolvedStyle {
  const styled = cell as ExcelJS.Cell & StyleCell;
  const fill = styled.fill;
  const fillColor =
    fill?.type === 'pattern' && fill.pattern === 'solid'
      ? resolveExcelColor(fill.fgColor, themeColors)
      : undefined;
  return {
    styleId: styleId ?? styled.styleId,
    fillColor,
    fontColor: resolveExcelColor(styled.font?.color, themeColors),
    italic: styled.font?.italic === true,
    bold: styled.font?.bold === true,
  };
}

function rgb(hex: string): [number, number, number] {
  const value = hex.replace('#', '');
  return [
    Number.parseInt(value.slice(0, 2), 16),
    Number.parseInt(value.slice(2, 4), 16),
    Number.parseInt(value.slice(4, 6), 16),
  ];
}

export function colorDistance(first: string, second: string): number {
  const a = rgb(first);
  const b = rgb(second);
  return Math.sqrt(
    (a[0] - b[0]) ** 2 * 0.3 + (a[1] - b[1]) ** 2 * 0.59 + (a[2] - b[2]) ** 2 * 0.11,
  );
}

export function isGreen(color?: string): boolean {
  if (!color) return false;
  const [red, green, blue] = rgb(color);
  return green > red * 1.12 && green > blue * 1.08;
}

export function isBlue(color?: string): boolean {
  if (!color) return false;
  const [red, green, blue] = rgb(color);
  return blue > red * 1.08 && blue >= green * 0.95;
}

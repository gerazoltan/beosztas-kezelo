import type ExcelJS from 'exceljs';
import type { FillCategory, ResolvedStyle } from '../domain/types';

interface ExcelColor {
  argb?: string;
  rgb?: string;
  theme?: number;
  tint?: number;
  indexed?: number;
}

interface RuntimeFill {
  type?: string;
  pattern?: string;
  fgColor?: ExcelColor;
  bgColor?: ExcelColor;
}

const INDEXED_COLORS = [
  '#000000',
  '#FFFFFF',
  '#FF0000',
  '#00FF00',
  '#0000FF',
  '#FFFF00',
  '#FF00FF',
  '#00FFFF',
  '#000000',
  '#FFFFFF',
  '#FF0000',
  '#00FF00',
  '#0000FF',
  '#FFFF00',
  '#FF00FF',
  '#00FFFF',
  '#800000',
  '#008000',
  '#000080',
  '#808000',
  '#800080',
  '#008080',
  '#C0C0C0',
  '#808080',
  '#9999FF',
  '#993366',
  '#FFFFCC',
  '#CCFFFF',
  '#660066',
  '#FF8080',
  '#0066CC',
  '#CCCCFF',
  '#000080',
  '#FF00FF',
  '#FFFF00',
  '#00FFFF',
  '#800080',
  '#800000',
  '#008080',
  '#0000FF',
  '#00CCFF',
  '#CCFFFF',
  '#CCFFCC',
  '#FFFF99',
  '#99CCFF',
  '#FF99CC',
  '#CC99FF',
  '#FFCC99',
  '#3366FF',
  '#33CCCC',
  '#99CC00',
  '#FFCC00',
  '#FF9900',
  '#FF6600',
  '#666699',
  '#969696',
  '#003366',
  '#339966',
  '#003300',
  '#333300',
  '#993300',
  '#993366',
  '#333399',
  '#333333',
] as const;

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
  const base = color.argb
    ? normalizeArgb(color.argb)
    : color.rgb
      ? normalizeArgb(color.rgb)
    : typeof color.theme === 'number'
      ? themeColors[color.theme]
      : typeof color.indexed === 'number'
        ? INDEXED_COLORS[color.indexed]
        : undefined;
  if (!base) return undefined;
  return color.tint ? tintColor(base, color.tint) : base;
}

export function describeExcelColor(color: ExcelColor | undefined): string | undefined {
  if (!color) return undefined;
  const parts: string[] = [];
  if (color.argb) parts.push(`argb=${color.argb.toUpperCase()}`);
  if (color.rgb) parts.push(`rgb=${color.rgb.toUpperCase()}`);
  if (typeof color.theme === 'number') parts.push(`theme=${color.theme}`);
  if (typeof color.indexed === 'number') parts.push(`indexed=${color.indexed}`);
  if (typeof color.tint === 'number') parts.push(`tint=${color.tint}`);
  return parts.length > 0 ? parts.join(', ') : undefined;
}

interface StyleCell {
  styleId?: number;
  fill?: ExcelJS.Fill;
  font?: Partial<ExcelJS.Font>;
}

export function fillIsVisible(fill: ExcelJS.Fill | undefined): boolean {
  if (!fill) return false;
  const runtime = fill as ExcelJS.Fill & RuntimeFill;
  if (runtime.type === 'pattern') {
    return typeof runtime.pattern === 'string' && runtime.pattern !== 'none';
  }
  return runtime.type === 'gradient';
}

function baseFillCategory(hasVisibleFill: boolean, fillColor: string | undefined): FillCategory {
  if (!hasVisibleFill) return 'noFill';
  if (isWhite(fillColor)) return 'white';
  if (isBlue(fillColor)) return 'blue';
  if (isGreen(fillColor)) return 'green';
  return 'unsupported';
}

export function resolveCellStyle(
  cell: ExcelJS.Cell,
  themeColors: string[],
  styleId?: number,
): ResolvedStyle {
  const styled = cell as ExcelJS.Cell & StyleCell;
  const fill = styled.fill;
  const runtimeFill = fill as (ExcelJS.Fill & RuntimeFill) | undefined;
  const hasVisibleFill = fillIsVisible(fill);
  const fillColor =
    hasVisibleFill && runtimeFill?.type === 'pattern' && runtimeFill.pattern === 'solid'
      ? resolveExcelColor(runtimeFill.fgColor, themeColors)
      : undefined;
  const rawFontColor = styled.font?.color;
  const normalizedFontColor = resolveExcelColor(rawFontColor, themeColors);
  const underline = styled.font?.underline;
  return {
    styleId: styleId ?? styled.styleId,
    fillType: runtimeFill?.type,
    fillPatternType: runtimeFill?.pattern,
    fillForegroundRaw: describeExcelColor(runtimeFill?.fgColor),
    fillBackgroundRaw: describeExcelColor(runtimeFill?.bgColor),
    fillColor,
    hasVisibleFill,
    fillCategory: baseFillCategory(hasVisibleFill, fillColor),
    fontColorRaw: describeExcelColor(rawFontColor),
    fontColor: normalizedFontColor ?? (rawFontColor ? undefined : '#000000'),
    underline: underline === true || (typeof underline === 'string' && underline !== 'none'),
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

export function isWhite(color?: string): boolean {
  return color?.toUpperCase() === '#FFFFFF';
}

export function isRed(color?: string): boolean {
  if (!color) return false;
  const [red, green, blue] = rgb(color);
  return red > green * 1.45 && red > blue * 1.45;
}

export function isBlack(color?: string): boolean {
  return color?.toUpperCase() === '#000000';
}
